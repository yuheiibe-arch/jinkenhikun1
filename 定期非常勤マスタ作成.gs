/**
 * =================================================================
 * 定期非常勤マスタ作成スクリプト（最終完成版・パーサー修正済み）
 * =================================================================
/**
 * =================================================================
 * 定期非常勤マスタ作成スクリプト（最終完成版・パーサー修正済み）
 * =================================================================
 */
function createPartTimeMaster() {
  const SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1miSCaOe0vZhWFqFHXBj6rvspebRi-53L0SzawjKgWm0/";
  const SOURCE_SHEET_NAME = "2025年度：医師情報";
  const OUTPUT_SHEET_NAME = "定期非常勤マスタ";
  const HOLIDAY_SHEET_NAME = "祝日";
  const TARGET_YEAR = 2025;

  try {
    Logger.log("--- 定期非常勤マスタの作成処理を開始します ---");
    SpreadsheetApp.getActiveSpreadsheet().toast("処理を開始しました...");

    const sourceSpreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
    const sourceSheet = sourceSpreadsheet.getSheetByName(SOURCE_SHEET_NAME);
    if (!sourceSheet) throw new Error(`シート「${SOURCE_SHEET_NAME}」が見つかりません。`);

    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const holidayList = getHolidaySet(activeSpreadsheet.getSheetByName(HOLIDAY_SHEET_NAME));
    
    const allData = sourceSheet.getDataRange().getValues();
    const header = allData[4];
    const dataRows = allData.slice(5);

    const colIdx = {
      status: header.indexOf("在籍区分"),
      doctorType: header.indexOf("医師区分"),
      licenseNum: header.indexOf("医籍番号"),
      name: header.indexOf("氏名\nスペース\nいれない！！"),
      specialty: header.indexOf("診療科"),
      contract: header.indexOf("2025年度契約内容"),
      startDate: header.indexOf("入職日"),
      holidayWork: header.indexOf("祝日"),
      newYearWork: header.indexOf("年末年始")
    };
    if (Object.values(colIdx).some(idx => idx === -1)) {
      const missing = Object.keys(colIdx).filter(key => colIdx[key] === -1);
      throw new Error(`ソースシートに必要なヘッダーが見つかりません: ${missing.join(', ')}`);
    }

    const finalOutputData = [];
    
    for (const row of dataRows) {
      const status = row[colIdx.status];
      const doctorType = row[colIdx.doctorType];

      if (doctorType === "定期非常勤" && (status === "01_在籍" || status === "99_入職前")) {
        const doctorName = row[colIdx.name] || "（氏名不明）";
        const contractText = row[colIdx.contract];
        if (!contractText) continue;

        const startDateValue = row[colIdx.startDate];
        const startDate = (startDateValue && startDateValue instanceof Date) ? new Date(Date.UTC(startDateValue.getFullYear(), startDateValue.getMonth(), startDateValue.getDate())) : null;
        
        // ★ 変更点: 呼び出す関数名を変更
        const scheduleRules = _masterCreator_parseContractToRules(contractText, TARGET_YEAR);
        const wageRules = parseWageRules(contractText);
        
        const hoursResult = calculateMonthlyHours(scheduleRules, TARGET_YEAR, holidayList, row[colIdx.holidayWork] !== "無", row[colIdx.newYearWork] !== "無", startDate);
        const averageWage = calculateAverageWage(wageRules);
        const totalWages = Math.round(hoursResult.totalHours * averageWage);

        const workplaces = extractWorkplaces(contractText);
        const wageDisplayText = createWageDisplayText(wageRules);
        const totalTimeAndDaysText = `${hoursResult.totalHours}h / ${hoursResult.totalDays}d`;
        
        finalOutputData.push([
          null, row[colIdx.licenseNum], doctorName, workplaces, contractText,
          row[colIdx.startDate], row[colIdx.specialty], wageDisplayText,
          status, row[colIdx.holidayWork], row[colIdx.newYearWork], TARGET_YEAR,
          ...hoursResult.monthlyHours, totalTimeAndDaysText,
          averageWage, totalWages
        ]);
      }
    }

    if (finalOutputData.length > 0) {
        const outputSheet = activeSpreadsheet.getSheetByName(OUTPUT_SHEET_NAME) || activeSpreadsheet.insertSheet(OUTPUT_SHEET_NAME);
        writeToPartTimeMasterSheet(outputSheet, finalOutputData);
        Logger.log(`--- 処理完了: ${finalOutputData.length}件のデータを「${OUTPUT_SHEET_NAME}」に出力しました ---`);
    } else {
        Logger.log("処理対象のデータが見つかりませんでした。");
    }
    SpreadsheetApp.getActiveSpreadsheet().toast("処理が完了しました。");
  } catch (e) {
    Logger.log(`エラーが発生しました: ${e.message}\nStack: ${e.stack}`);
    SpreadsheetApp.getUi().alert(`処理中にエラーが発生しました:\n\n${e.message}`);
  }
}

// =================================================================
// ヘルパー関数群（最終修正版）
// =================================================================

/**
 * 【最終版】契約テキストを解析し、勤務ルールの配列を生成します。
 */
// ★ 変更点: 関数名を変更
function _masterCreator_parseContractToRules(text, year) {
  if (!text || typeof text !== 'string') return [];
  const lines = text.split('\n').map(line => line.trim()).filter(line => line);
  const rules = [];
  const dayMap = { "日": 0, "月": 1, "火": 2, "水": 3, "木": 4, "金": 5, "土": 6 };

  const parseLine = (line, start, end) => {
    if (line.startsWith('(') && line.endsWith(')')) return;
    const dayMatch = line.match(/(月|火|水|木|金|土|日)曜日?/);
    if (!dayMatch) return;
    const dayOfWeek = dayMap[dayMatch[0].charAt(0)];
    const weekNumMatch = line.match(/第([0-9・,./\s]+)/);
    const targetWeeks = weekNumMatch ? weekNumMatch[1].split(/[・,./\s]/).filter(w => w && !isNaN(w)).map(Number) : null;
    const primaryTimeMatch = line.match(/(\d{1,2}(?::\d{2})?)\s*[-‐〜～~]\s*(\d{1,2}(?::\d{2})?)/);
    if (!primaryTimeMatch) return;
    let shiftStart = new Date(`1970/01/01 ${normalizeTime(primaryTimeMatch[1])}`);
    let shiftEnd = new Date(`1970/01/01 ${normalizeTime(primaryTimeMatch[2])}`);
    let hours = (shiftEnd - shiftStart) / (1000 * 60 * 60);
    if (isNaN(hours)) return;
    const breakMatch = line.match(/\(休憩\s*(\d{1,2}:\d{2})[\s]*[～\-~‐〜][\s]*(\d{1,2}:\d{2})\s*\)/);
    if (breakMatch) {
      let breakHours = (new Date(`1970/01/01 ${normalizeTime(breakMatch[2])}`) - new Date(`1970/01/01 ${normalizeTime(breakMatch[1])}`)) / (1000 * 60 * 60);
      hours -= breakHours;
    } else if (!line.includes("休憩時間なし")) {
      const lunchStart = new Date(`1970/01/01 13:00`);
      const lunchEnd = new Date(`1970/01/01 15:00`);
      if(shiftStart < lunchEnd && shiftEnd > lunchStart && hours > 4) { hours -= 2; }
    }
    if(hours > 0) rules.push({ start, end, dayOfWeek, hours, weeks: targetWeeks });
  };
  const periodMarkers = [];
  lines.forEach((line, index) => {
    let period = { start: null, end: null };
    const explicitRangeMatch = line.match(/(\d{4})\/(\d{1,2})\/(\d{1,2})\s*[~～〜‐-]\s*(\d{4})\/(\d{1,2})\/(\d{1,2})/);
    const startOnlyMatch = line.match(/^(?:(\d{4})年)?(\d{1,2})\s*月\s*(\d{1,2})?\s*[日]?\s*[~～〜‐-]/);
    if (explicitRangeMatch) {
      period.start = new Date(Date.UTC(parseInt(explicitRangeMatch[1]), parseInt(explicitRangeMatch[2]) - 1, parseInt(explicitRangeMatch[3])));
      period.end = new Date(Date.UTC(parseInt(explicitRangeMatch[4]), parseInt(explicitRangeMatch[5]) - 1, parseInt(explicitRangeMatch[6])));
    } else if (startOnlyMatch) {
      let startYear = startOnlyMatch[1] ? parseInt(startOnlyMatch[1], 10) : (parseInt(startOnlyMatch[2], 10) < 4 ? year + 1 : year);
      let startDay = startOnlyMatch[3] ? parseInt(startOnlyMatch[3], 10) : 1;
      period.start = new Date(Date.UTC(startYear, parseInt(startOnlyMatch[2], 10) - 1, startDay));
    }
    if (period.start) periodMarkers.push({ index, period });
  });
  const fiscalStart = new Date(Date.UTC(year, 3, 1));
  const fiscalEnd = new Date(Date.UTC(year + 1, 2, 31));
  if (periodMarkers.length === 0) {
    lines.forEach(line => parseLine(line, fiscalStart, fiscalEnd));
  } else {
    let lastIndex = -1;
    let currentStart = fiscalStart;
    periodMarkers.forEach(marker => {
      const blockLines = lines.slice(lastIndex + 1, marker.index);
      const blockEnd = new Date(marker.period.start.getTime() - 86400000);
      blockLines.forEach(line => parseLine(line, currentStart, blockEnd));
      currentStart = marker.period.start;
      lastIndex = marker.index;
    });
    const finalBlockLines = lines.slice(lastIndex + 1);
    const finalEnd = periodMarkers[periodMarkers.length - 1].period.end || fiscalEnd;
    finalBlockLines.forEach(line => parseLine(line, currentStart, finalEnd));
  }
  return rules;
}

function calculateMonthlyHours(scheduleRules, year, holidayList, hasHolidayWork, hasNewYearWork, employmentStartDate) {
  const monthlyHours = Array(12).fill(0);
  const workDays = new Set();
  const fiscalStart = new Date(Date.UTC(year, 3, 1));
  const fiscalEnd = new Date(Date.UTC(year + 1, 3, 0));
  for (let d = new Date(fiscalStart.getTime()); d <= fiscalEnd; d.setUTCDate(d.getUTCDate() + 1)) {
    if (employmentStartDate && d < employmentStartDate) continue;
    const month = d.getUTCMonth() + 1;
    const day = d.getUTCDate();
    const isNewYearHoliday = (month === 12 && day >= 29) || (month === 1 && day <= 3);
    if (isNewYearHoliday && !hasNewYearWork) continue;
    const dateStr = `${d.getUTCFullYear()}-${month}-${day}`;
    if (holidayList.has(dateStr) && !hasHolidayWork) continue;
    const dayOfWeek = d.getUTCDay();
    let isWorkDay = false;
    for (const scheduleRule of scheduleRules) {
      if (scheduleRule.dayOfWeek === dayOfWeek && d >= scheduleRule.start && d <= scheduleRule.end) {
        if (scheduleRule.weeks && !scheduleRule.weeks.includes(Math.floor((d.getUTCDate() - 1) / 7) + 1)) continue;
        const monthIndex = d.getUTCMonth() < 3 ? d.getUTCMonth() + 9 : d.getUTCMonth() - 3;
        monthlyHours[monthIndex] += scheduleRule.hours;
        isWorkDay = true;
      }
    }
    if (isWorkDay) workDays.add(dateStr);
  }
  const totalHours = monthlyHours.reduce((a, b) => a + b, 0);
  return {
    monthlyHours: monthlyHours.map(h => Math.round(h * 10) / 10),
    totalHours: Math.round(totalHours * 10) / 10,
    totalDays: workDays.size
  };
}
function calculateAverageWage(wageRules) {
  let totalWageProduct = 0;
  let totalHours = 0;
  const rulesToConsider = wageRules.normal.length > 0 ? wageRules.normal : wageRules.holiday;
  if (rulesToConsider.length === 0) return 0;
  const uniqueRules = [];
  const seen = new Set();
  rulesToConsider.forEach(rule => {
    const key = `${rule.start}-${rule.end}-${rule.wage}`;
    if (!seen.has(key)) {
      uniqueRules.push(rule);
      seen.add(key);
    }
  });
  uniqueRules.forEach(rule => {
      const start = new Date(`1970-01-01T${rule.start}:00Z`);
      const end = new Date(`1970-01-01T${rule.end}:00Z`);
      let hours = (end - start) / (1000 * 60 * 60);
      if (hours > 0) {
          const lunchStart = new Date(`1970-01-01T13:00:00Z`);
          const lunchEnd = new Date(`1970-01-01T15:00:00Z`);
          if(start < lunchEnd && end > lunchStart) {
              hours -= 2;
          }
          if (hours > 0) {
            totalWageProduct += hours * rule.wage;
            totalHours += hours;
          }
      }
  });
  return totalHours > 0 ? Math.round(totalWageProduct / totalHours) : 0;
}
function writeToPartTimeMasterSheet(sheet, data) {
  const header = [
    "番号", "医籍番号", "医師名", "主勤務先", "契約内容", "入職日", "専門", "時給",
    "在籍区分", "祝日勤務", "年末年始勤務", "年度",
    "4月", "5月", "6月", "7月", "8月", "9月",
    "10月", "11月", "12月", "1月", "2月", "3月",
    "延べ時間/日数", "平均時給", "年間賃金見込み"
  ];
  sheet.clear();
  const dataRange = sheet.getRange(2, 1, data.length, data[0].length);
  sheet.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight("bold");
  if (data && data.length > 0) {
    data.forEach((row, index) => row[0] = index + 1);
    dataRange.setValues(data);
    dataRange.setHorizontalAlignment('left').setVerticalAlignment('top').setWrap(true);
    sheet.getRange(2, 1, data.length, 1).setHorizontalAlignment('center'); 
    sheet.getRange(2, 13, data.length, 15).setHorizontalAlignment('center');
    sheet.getRange(1, 1, 1, header.length).setHorizontalAlignment('center');
  }
}
function normalizeTime(timeStr) {
  if (!timeStr) return "00:00";
  timeStr = timeStr.replace(/：/g, ':');
  if (!timeStr.includes(':')) { return timeStr.padStart(2, '0') + ':00'; }
  const parts = timeStr.split(':');
  return parts[0].padStart(2, '0') + ':' + parts[1].padEnd(2, '0');
};
function parseWageRules(contractText) {
  if (!contractText || typeof contractText !== 'string') { return { normal: [], holiday: [] }; }
  const extractRulesFromText = (text) => {
    const rules = [];
    const complexRegex = /(\d{1,2}(?::\d{2})?)\s*[-‐〜～~]\s*(\d{1,2}(?::\d{2})?)[^円\d]*?([\d,]+)円/g;
    const complexMatches = [...text.matchAll(complexRegex)];
    if (complexMatches.length > 0) {
      for (const match of complexMatches) {
        rules.push({
          start: normalizeTime(match[1]),
          end: normalizeTime(match[2]),
          wage: parseInt(match[3].replace(/,/g, ''), 10)
        });
      }
      return rules;
    }
    const simpleRegex = /(?:時給|平日時給|待機時給)?\s*[:：]?\s*([\d,]+)円/;
    const simpleMatch = text.match(simpleRegex);
    if (simpleMatch) {
      rules.push({ start: "00:00", end: "23:59", wage: parseInt(simpleMatch[1].replace(/,/g, ''), 10) });
    }
    return rules;
  };
  let normalText = contractText;
  let holidayText = "";
  const holidaySplitMatch = contractText.match(/(.*?)(\(祝日|祝日時給|祝日：)(.*)/s);
  if (holidaySplitMatch) {
    normalText = holidaySplitMatch[1];
    holidayText = holidaySplitMatch[2] + holidaySplitMatch[3];
  }
  let normalRules = extractRulesFromText(normalText);
  let holidayRules = extractRulesFromText(holidayText);
  if (holidayRules.length === 0 && holidaySplitMatch) {
      const simpleHolidayRegex = /([\d,]+)円/;
      const simpleHolidayMatch = holidayText.match(simpleHolidayRegex);
      if(simpleHolidayMatch){
        holidayRules.push({ start: "00:00", end: "23:59", wage: parseInt(simpleHolidayMatch[1].replace(/,/g, ''), 10) });
      }
  }
  if (holidayRules.length === 0 && normalRules.length > 0) {
    holidayRules = JSON.parse(JSON.stringify(normalRules));
  }
  if (normalRules.length === 0 && holidayRules.length > 0) {
    normalRules = JSON.parse(JSON.stringify(holidayRules));
  }
  return { normal: normalRules, holiday: holidayRules };
}
function getHolidaySet(holidaySheet) {
  if (!holidaySheet) return new Set();
  const values = holidaySheet.getRange("A2:A" + holidaySheet.getLastRow()).getValues();
  const holidaySet = new Set();
  values.forEach(row => {
    if (row[0] instanceof Date) {
      holidaySet.add(`${row[0].getFullYear()}-${row[0].getMonth() + 1}-${row[0].getDate()}`);
    }
  });
  return holidaySet;
}
function extractWorkplaces(contractText) {
  if (!contractText) return "";
  const regex = /【([^】]+)】/g;
  const matches = [...contractText.matchAll(regex)];
  if (matches.length === 0) return "";
  const workplaces = new Set(matches.map(m => m[1].trim()));
  return Array.from(workplaces).join(', ');
}
function createWageDisplayText(wageRules) {
  const formatUniqueRules = (rules) => {
    if (!rules || rules.length === 0) return "N/A";
    const uniqueRules = [];
    const seen = new Set();
    rules.forEach(rule => {
      const key = `${rule.start}-${rule.end}-${rule.wage}`;
      if (!seen.has(key)) {
        uniqueRules.push(rule);
        seen.add(key);
      }
    });
    return uniqueRules.map(r => `${r.wage.toLocaleString()}円(${r.start}-${r.end})`).join(' ');
  };
  const normalWageStr = `平日: ${formatUniqueRules(wageRules.normal)}`;
  const holidayWageStr = `祝日: ${formatUniqueRules(wageRules.holiday)}`;
  return `${normalWageStr} | ${holidayWageStr}`;
}