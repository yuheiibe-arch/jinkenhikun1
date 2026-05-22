// =================================================================
// ■ 計算エンジン (休憩ロジックを元に戻した最終修正版)
// =================================================================

function calculateMonthlyHoursFromContract(contractText, year, hasHolidayWork, holidayList, hasNewYearWork) {
  // ★ 変更点: 呼び出す関数名を変更
  const scheduleRules = _calcEngine_parseContractToRules(contractText, year);
  
  const monthlyHours = Array(12).fill(0);
  const monthlyWorkDays = Array(12).fill(0);
  let totalWorkDays = 0;

  if (scheduleRules.length === 0) {
    return { monthlyHours, totalWorkDays: 0, monthlyWorkDays };
  }
  
  const fiscalStart = new Date(Date.UTC(year, 3, 1));
  const fiscalEnd = new Date(Date.UTC(year + 1, 3, 0));
  const workDaysCount = new Set();

  for (let d = new Date(fiscalStart.getTime()); d <= fiscalEnd; d.setUTCDate(d.getUTCDate() + 1)) {
    const dayOfWeek = d.getUTCDay();
    const dateStr = `${d.getUTCFullYear()}-${d.getUTCMonth() + 1}-${d.getUTCDate()}`;

    const month = d.getUTCMonth() + 1;
    const day = d.getUTCDate();
    const isNewYearHoliday = (month === 12 && day >= 29) || (month === 1 && day <= 3);

    if (isNewYearHoliday && !hasNewYearWork) continue;
    
    for (const rule of scheduleRules) {
      if (rule.dayOfWeek === dayOfWeek && d >= rule.start && d <= rule.end) {
        if (rule.weeks && !rule.weeks.includes(Math.floor((d.getUTCDate() - 1) / 7) + 1)) continue;
        if (holidayList.has(dateStr) && !hasHolidayWork) continue;

        const monthIndex = d.getUTCMonth() < 3 ? d.getUTCMonth() + 9 : d.getUTCMonth() - 3;
        
        if (!workDaysCount.has(dateStr)) {
          monthlyWorkDays[monthIndex]++;
        }

        monthlyHours[monthIndex] += rule.hours;
        workDaysCount.add(dateStr);
      }
    }
  }

  totalWorkDays = workDaysCount.size;
  const roundedMonthlyHours = monthlyHours.map(h => Math.round(h * 100) / 100);
  
  return { monthlyHours: roundedMonthlyHours, totalWorkDays, monthlyWorkDays };
}

// ★ 変更点: 関数名を変更
function _calcEngine_parseContractToRules(text, year) {
  if (!text || typeof text !== 'string') return [];

  const lines = text.split('\n').map(line => line.trim()).filter(line => line);
  const rules = [];
  const dayMap = { "日": 0, "月": 1, "火": 2, "水": 3, "木": 4, "金": 5, "土": 6 };

  const parseLineForRule = (line, start, end) => {
    const dayMatch = line.match(/(月|火|水|木|金|土|日)曜日?/);
    if (!dayMatch) return;
    const dayOfWeek = dayMap[dayMatch[1]];
    const weekNumMatch = line.match(/第([0-9・,/\s]+)/);
    const targetWeeks = weekNumMatch ? weekNumMatch[1].split(/[・,/\s]/).filter(w => w).map(Number) : null;
    
    let totalHours = 0;
    
    const timeRegex = /(\d{1,2}:\d{2})[\s]*[～\-~‐〜][\s]*(\d{1,2}:\d{2})/g;
    const timeMatches = [...line.matchAll(timeRegex)];
    if (timeMatches.length > 0) {
      let mainStart, mainEnd;
      timeMatches.forEach((match, index) => {
        const s = new Date(`1970/01/01 ${match[1]}`);
        const e = new Date(`1970/01/01 ${match[2]}`);
        if (index === 0) {
          mainStart = s;
          mainEnd = e;
        }
        totalHours += (e - s) / (1000 * 60 * 60);
      });

      if (timeMatches.length === 1 && totalHours > 4) {
        const lunchStart = new Date(`1970/01/01 13:00`);
        const lunchEnd = new Date(`1970/01/01 15:00`);
        if (mainStart < lunchEnd && mainEnd > lunchStart && mainStart.getHours() < 13) {
          totalHours -= 2;
        }
      }
    }
        
    if (totalHours > 0) {
      rules.push({ start, end, dayOfWeek, hours: totalHours, weeks: targetWeeks });
    }
  };

  const periodMarkers = [];
  lines.forEach((line, index) => {
    let match;
    if (match = line.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})\s*[~～]\s*(\d{4})\/(\d{1,2})\/(\d{1,2})/)) {
      const period = {
        start: new Date(Date.UTC(parseInt(match[1]), parseInt(match[2]) - 1, parseInt(match[3]))),
        end: new Date(Date.UTC(parseInt(match[4]), parseInt(match[5]) - 1, parseInt(match[6])))
      };
      periodMarkers.push({ index, period });
      return;
    }
    if (match = line.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})\s*[~～]/)) {
      const period = {
        start: new Date(Date.UTC(parseInt(match[1]), parseInt(match[2]) - 1, parseInt(match[3]))),
        end: null
      };
      periodMarkers.push({ index, period });
    }
  });

  const fiscalStart = new Date(Date.UTC(year, 3, 1));
  const fiscalEnd = new Date(Date.UTC(year + 1, 2, 31));

  if (periodMarkers.length === 0) {
    lines.forEach(line => parseLineForRule(line, fiscalStart, fiscalEnd));
  } else {
    periodMarkers.forEach((marker, i) => {
      const startLine = marker.index + 1;
      const endLine = (i + 1 < periodMarkers.length) ? periodMarkers[i + 1].index : lines.length;
      const blockLines = lines.slice(startLine, endLine);
      const startDate = marker.period.start;
      const endDate = marker.period.end || ((i + 1 < periodMarkers.length) ? new Date(periodMarkers[i + 1].period.start.getTime() - 86400000) : fiscalEnd);
      blockLines.forEach(line => parseLineForRule(line, startDate, endDate));
    });
  }
  
  return rules;
}