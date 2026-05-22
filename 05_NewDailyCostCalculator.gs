/**
 * ====================================================================
 * 【ファイル: 05_NewDailyCostCalculator.gs】(完全一致修正版)
 * 日次コスト計算の統合・分岐エンジン（ルーター）
 * ====================================================================
 */

const NewDailyCostCalculator = (function() {
  
  const NEW_SYSTEM_START_DATE = new Date('2026-04-01T00:00:00+09:00');
  const AGENCY_FEE_RATE = 0.20; 

  const IGNORE_CLINIC_KEYWORDS = [
    "有給", "欠勤", "院外勤務", "医師会業務", "バックアップシフト", "嘱託医業務"
  ];

  function buildMonthlyCache(startDate, endDate) {
    if (startDate < NEW_SYSTEM_START_DATE && endDate < NEW_SYSTEM_START_DATE) {
      return null; 
    }

    const cache = {
      shifts: NewDataFetcher.fetchShiftData(startDate, endDate),
      ftDoctors: new Set(), 
      ptDoctors: new Set()  
    };

    try {
      const ftData = NewDataFetcher.fetchFullTimeMaster2026();
      for (let i = 1; i < ftData.length; i++) {
        const medId = String(ftData[i][1]).trim();
        if (medId) cache.ftDoctors.add(medId);
      }
    } catch (e) {
      Logger.log(`⚠️ 常勤マスタ(2026)の取得をスキップ: ${e.message}`);
    }

    try {
      const ptData = NewDataFetcher.fetchPartTimeMaster2026();
      for (let i = 1; i < ptData.length; i++) {
        const medId = String(ptData[i][1]).trim();
        if (medId) cache.ptDoctors.add(medId);
      }
    } catch (e) {
      Logger.log(`⚠️ 定期非常勤マスタ(2026)の取得をスキップ: ${e.message}`);
    }

    return cache;
  }

  function routeAndCalculateDailyCost(targetDate, allDailyData, settings, cache) {
    if (!(targetDate instanceof Date) || isNaN(targetDate.getTime())) {
      throw new Error("【システムエラー】無効な日付が渡されました。");
    }

    if (targetDate < NEW_SYSTEM_START_DATE) {
      if (typeof _calculateDailyPersonnelCosts === 'function') {
        _calculateDailyPersonnelCosts(targetDate, allDailyData, settings);
      }
    } else {
      if (!cache) throw new Error("【システムエラー】キャッシュオブジェクトが必要です。");
      _processNewDailyCost(targetDate, allDailyData, settings, cache);
    }
  }

  function _processNewDailyCost(targetDate, allDailyData, settings, cache) {
    const y = targetDate.getFullYear();
    const m = targetDate.getMonth() + 1; // ゼロ埋めなし
    const d = targetDate.getDate();      // ゼロ埋めなし
    
    // ★ 旧システムと完全に一致するキー (例: 2026/4/1)
    const dateKey = `${y}/${m}/${d}`;

    // JST文字列で日付を確保
    const fetchDateStr = Utilities.formatDate(targetDate, "JST", "yyyy/MM/dd");
    const holDateStr = Utilities.formatDate(targetDate, "JST", "yyyy-MM-dd");

    if (!allDailyData[dateKey]) allDailyData[dateKey] = {};

    const isHol = settings.holidays ? settings.holidays.has(holDateStr) : false;
    const dailyShifts = cache.shifts.filter(shift => shift.workDateStr === fetchDateStr);

    dailyShifts.forEach(shift => {
      if (!shift.startTimeStr || !shift.endTimeStr || !shift.clinicName) return;

      const rawClinic = String(shift.clinicName);
      
      if (IGNORE_CLINIC_KEYWORDS.some(kw => rawClinic.includes(kw))) {
        return; 
      }

      let clinic = rawClinic;
      if (typeof gt_normalizeClinicName === 'function' && settings.clinicAliasMap) {
        clinic = gt_normalizeClinicName(rawClinic, settings.clinicAliasMap);
      }

      const medId = String(shift.medId).trim();
      
      if (!allDailyData[dateKey][clinic]) {
        // ★ 旧システムと完全に一致するプロパティ名 (totalCost, totalHours)
        allDailyData[dateKey][clinic] = {
          totalCost: 0, 
          totalHours: 0, 
          sales: 0,
          costDetails: {
            '常勤医師給与': 0, '定期非常勤給与': 0, 'スポット医師給与': 0,
            '直接応募医師給与': 0, '紹介会社医師給与': 0, '所定休出医師給与': 0,
            '紹介手数料': 0, '特別手当': 0
          }
        };
      }

      const dailyClinicObj = allDailyData[dateKey][clinic];
      const comment = shift.comment1 || "";
      const isAgency = comment.includes("紹介会社");
      const isKyushutsu = comment.includes("所定休出");

      let category = "";
      if (isKyushutsu) {
        category = "所定休出医師給与";
      } else if (cache.ftDoctors.has(medId)) {
        category = "常勤医師給与";
      } else if (cache.ptDoctors.has(medId)) {
        category = "定期非常勤給与";
      } else if (isAgency) {
        category = "紹介会社医師給与";
      } else {
        category = "直接応募医師給与";
      }

      let dailyWageSum = 0;
      let dailyHoursSum = 0;

      try {
        const slots = NewWageEngine.calculate(
          medId, rawClinic, shift.department, targetDate, 
          shift.startTimeStr, shift.endTimeStr, isHol
        );

        slots.forEach(slot => {
          const startH = parseInt(slot.startTime.split(':')[0], 10);
          const startM = parseInt(slot.startTime.split(':')[1], 10) / 60;
          const endH = parseInt(slot.endTime.split(':')[0], 10);
          const endM = parseInt(slot.endTime.split(':')[1], 10) / 60;
          
          let duration = (endH + endM) - (startH + startM);
          if (duration < 0) duration += 24; 
          
          if (duration > 0) {
            dailyHoursSum += duration;
            dailyWageSum += slot.wage * duration;
          }
        });
      } catch (e) {
        Logger.log(`⚠️ ${dateKey} [${rawClinic}] ${shift.doctorName} の計算エラー: ${e.message}`);
        return; 
      }

      let extraAllowance = 0; 
      
      // ★ 正しいプロパティ名に加算
      dailyClinicObj.totalCost += dailyWageSum + extraAllowance;
      dailyClinicObj.totalHours += dailyHoursSum;
      dailyClinicObj.costDetails[category] += dailyWageSum;
      dailyClinicObj.costDetails['特別手当'] += extraAllowance;
      
      if (isAgency) {
        const agencyFee = dailyWageSum * AGENCY_FEE_RATE;
        dailyClinicObj.totalCost += agencyFee;
        dailyClinicObj.costDetails['紹介手数料'] += agencyFee;
      }
    });
  }

  return {
    buildMonthlyCache: buildMonthlyCache,
    routeAndCalculateDailyCost: routeAndCalculateDailyCost
  };
})();