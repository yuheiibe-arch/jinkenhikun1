/**
 * =================================================================
 * 【ファイル X/X: グラフ生成ツール・サーバー処理 (I_Graph_Tool.gs)】- 最終FIX版
 * =================================================================
 */

function gt_calculateHybridMonth(clinic, monthStr, settings) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const [year, month] = monthStr.split('/').map(Number);
  const daysInMonth = new Date(year, month, 0).getDate();
  const normalizedClinic = gt_normalizeClinicName(clinic, settings.clinicAliasMap);

  const monthlyTotal = {
    sales: 0, cost: 0, hours: 0, budget: 0,
    costDetails: { '常勤医師給与': 0, '定期非常勤給与': 0, '直接応募医師給与': 0, '紹介会社医師給与': 0, '所定休出医師給与': 0, '紹介手数料': 0, '特別手当': 0 },
    doctorCount: { '常勤': 0, '定期非常勤': 0, '直接応募医師': 0, '紹介会社医師': 0 }
  };
  
  const actualsData = _calculateActualsForAllClinics_OneMonth(monthStr, settings, today);
  const clinicActuals = actualsData.clinicDataMap.get(normalizedClinic);

  if (clinicActuals) {
    monthlyTotal.cost = clinicActuals.cost || 0;
    monthlyTotal.hours = clinicActuals.hours || 0;
    monthlyTotal.costDetails = clinicActuals.costDetails;
    monthlyTotal.doctorCount = clinicActuals.doctorCount;
  }

  for (let day = today.getDate(); day <= daysInMonth; day++) {
    const targetDate = new Date(year, month - 1, day);
    if (targetDate < today) continue;

    const dayProjection = gt_calculateFuturePersonnelCosts(clinic, monthStr, settings, targetDate);
    monthlyTotal.cost += dayProjection.totalCost || 0;
    monthlyTotal.hours += dayProjection.totalHours || 0;
    for (const key in monthlyTotal.costDetails) {
      monthlyTotal.costDetails[key] += dayProjection.costDetails[key] || 0;
    }
  }
  
  monthlyTotal.sales = gt_calculateSalesForMonth(clinic, monthStr, settings, false);
  const budgetData = settings.budgetMap.get(normalizedClinic) || {};
  monthlyTotal.budget = (budgetData.ft_budget?.[month] || 0) + (budgetData.pt_budget?.[month] || 0) + (budgetData.spot_budget?.[month] || 0) + (budgetData.second_exam_budget?.[month] || 0);
  monthlyTotal.averages = _calculateAveragesFromActualsMap(actualsData.clinicDataMap, settings.clinicGroupMap);

  return monthlyTotal;
}

function gt_startCalculation(params) {
  try {
    const year = parseInt(params.year, 10);
    const months = [];
    for (let i = 0; i < 12; i++) {
      const month = (i + 4 > 12) ? (i + 4 - 12) : (i + 4);
      const targetYear = (i + 4 > 12) ? year + 1 : year;
      months.push(`${targetYear}/${String(month).padStart(2, '0')}`);
    }
    
    const sessionData = {
      clinic: params.clinic,
      monthsToProcess: months,
      allMonthlyData: []
    };

    const cache = CacheService.getScriptCache();
    cache.put('graphtool_session', JSON.stringify(sessionData), 1800);

    return { success: true, totalMonths: months.length };
  } catch(e) {
    Logger.log(`gt_startCalculation Error: ${e.message}`);
    return { success: false, message: e.message };
  }
}

// 【I_Graph_Tool.gs】に貼り付け
function gt_processOneMonth(monthIndex) {
  const cache = CacheService.getScriptCache();
  const sessionJson = cache.get('graphtool_session');
  if (!sessionJson) return "エラー: セッションがタイムアウトしました。";

  const session = JSON.parse(sessionJson);
  const monthStr = session.monthsToProcess[monthIndex];
  
  try {
    const settings = _loadInitialData(SpreadsheetApp.getActiveSpreadsheet());
    
    const today = new Date();
    const todayJST = new Date(today.toLocaleString('en-US', { timeZone: 'Asia/Tokyo' }));
    todayJST.setHours(0, 0, 0, 0);
    const targetMonthDate = new Date(monthStr + "/01");
    
    let monthlyData;
    let type = '';

    if (targetMonthDate.getFullYear() < todayJST.getFullYear() || (targetMonthDate.getFullYear() === todayJST.getFullYear() && targetMonthDate.getMonth() < todayJST.getMonth())) {
      type = '実績';
      monthlyData = gt_calculateActualsForMonth(session.clinic, monthStr, settings);
    } 
    else if (targetMonthDate.getFullYear() === todayJST.getFullYear() && targetMonthDate.getMonth() === todayJST.getMonth()) {
      type = '実績/予測';
      monthlyData = gt_calculateHybridMonth(session.clinic, monthStr, settings);
    }
    else {
      type = '予測';
      const baseForecast = gt_calculateFuturePersonnelCosts(session.clinic, monthStr, settings);
      const actuals = session.allMonthlyData.filter(d => (d.type === '実績' || d.type === '実績/予測'));
      
      const averageCosts = { '直接応募医師給与': 0, '紹介会社医師給与': 0, '所定休出医師給与': 0, '紹介手数料': 0, '特別手当': 0 };
      if (actuals.length > 0) {
        for (const key in averageCosts) {
          const sum = actuals.reduce((acc, data) => acc + (data.costDetails[key] || 0), 0);
          averageCosts[key] = sum / actuals.length;
        }
      }
      
      const averageDoctorCount = { '直接応募医師': 0, '紹介会社医師': 0 };
      if (actuals.length > 0) {
        for (const key in averageDoctorCount) {
          const sum = actuals.reduce((acc, data) => acc + (data.doctorCount[key] || 0), 0);
          averageDoctorCount[key] = sum / actuals.length;
        }
      }

      const finalCostDetails = {
        ...baseForecast.costDetails,
        ...averageCosts
      };
      
      const totalCost = Object.values(finalCostDetails).reduce((acc, val) => acc + val, 0);
      
      const finalDoctorCount = {
        ...baseForecast.doctorCount,
        '直接応募医師': averageDoctorCount['直接応募医師'],
        '紹介会社医師': averageDoctorCount['紹介会社医師']
      };

      const sales = gt_calculateSalesForMonth(session.clinic, monthStr, settings, true);
      
      // ★★★ 修正箇所 ★★★
      // 過去月の予算を流用せず、予測月の正しい予算を計算する
      const normalizedClinic = gt_normalizeClinicName(session.clinic, settings.clinicAliasMap);
      const month = parseInt(monthStr.split('/')[1], 10);
      const budgetData = settings.budgetMap.get(normalizedClinic) || {};
      const totalBudget = (budgetData.ft_budget?.[month] || 0) + (budgetData.pt_budget?.[month] || 0) + (budgetData.spot_budget?.[month] || 0) + (budgetData.second_exam_budget?.[month] || 0);

      const lastActualMonthData = session.allMonthlyData.slice().reverse().find(d => d.type === '実績');

      monthlyData = {
        sales: sales,
        cost: totalCost,
        hours: baseForecast.totalHours,
        budget: totalBudget, // ★ 正しい予算をセット
        costDetails: finalCostDetails,
        doctorCount: finalDoctorCount,
        averages: (lastActualMonthData ? lastActualMonthData.averages : {})
      };
    }
    
    const dataForThisMonth = { month: monthStr, type: type, ...monthlyData };
    session.allMonthlyData.push(dataForThisMonth);
    cache.put('graphtool_session', JSON.stringify(session), 1800);

    return `${monthStr} の${type}データを計算しました。`;
  } catch(e) {
    Logger.log(`gt_processOneMonth Error (${monthStr}): ${e.message}\n${e.stack}`);
    return `エラー (${monthStr}): ${e.message}`;
  }
}

function gt_finishCalculation(params) {
  const cache = CacheService.getScriptCache();
  const sessionJson = cache.get('graphtool_session');
  if (!sessionJson) return "エラー: セッションがタイムアウトしました。";
  
  const session = JSON.parse(sessionJson);
  const settings = _loadInitialData(SpreadsheetApp.getActiveSpreadsheet());
  
  gt_writeToMonthlySummaryAndCreateChart(
    SpreadsheetApp.getActiveSpreadsheet(),
    params.clinic,
    session.allMonthlyData,
    params.charts,
    settings
  );
  
  gt_cancelCalculation();
  return "月間サマリーシートの作成とグラフの描画が完了しました。";
}

function gt_cancelCalculation() {
  const cache = CacheService.getScriptCache();
  cache.remove('graphtool_session');
}