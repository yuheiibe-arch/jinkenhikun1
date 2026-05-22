/**
 * =================================================================
 * 【ファイル差し替え: 当月コスト日次自動集計 (Z_AutoCalculator.gs)】
 * =================================================================
 * 備考: この関数を毎日実行するトリガーを設定することで、
 * 「拠点別コスト」シートを自動で最新の状態に保ちます。
 * 【2026年度新仕様対応版】NewDailyCostCalculatorによる統合ルーティング
 */

function runDailyCostCalculationForCurrentMonth() {
  const SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. 当月の年月を取得し、日付範囲を設定
  const today = new Date();
  const year = today.getFullYear();
  const month = today.getMonth() + 1; // getMonth()は0-11を返すため+1
  const monthStr = `${year}/${String(month).padStart(2, '0')}`;
  
  const startDate = new Date(year, month - 1, 1);
  const endDate = new Date(year, month, 0);
  const totalDaysInMonth = endDate.getDate();

  Logger.log(`自動集計を開始します。対象月: ${monthStr}`);

  try {
    // 2. 全ツール共通の設定データを読み込む
    const settings = _loadInitialData(SPREADSHEET);
    if (!settings) {
      throw new Error("マスターデータの読み込みに失敗しました。処理を中断します。");
    }
    Logger.log("共通設定の読み込み完了。");

    // 3. 昨年度同月の実績データを参照用に取得
    const lastYearMonthStr = `${year - 1}/${String(month).padStart(2, '0')}`;
    let lastYearData = {};
    try {
      if (typeof _calculateLastYearActuals_FOR_DAILY_TOOL === 'function') {
        lastYearData = _calculateLastYearActuals_FOR_DAILY_TOOL(lastYearMonthStr, settings);
      }
    } catch (e) {
      Logger.log(`警告: 昨年度 (${lastYearMonthStr}) の実績データ取得に失敗しました。${e.message}`);
    }
    settings.lastYearDailyData = lastYearData; 
    Logger.log("昨年度実績の読み込み完了。");

    // 4. メインとなる日次データ格納用オブジェクトを初期化
    const allDailyData = {};
    const todayForCompare = new Date();
    todayForCompare.setHours(0, 0, 0, 0);

    // ★ 新仕様: 月間データの一括キャッシュ（API呼び出しを最小化して超高速化）
    const monthlyCache = (typeof NewDailyCostCalculator !== 'undefined')
      ? NewDailyCostCalculator.buildMonthlyCache(startDate, endDate)
      : null;

    // 5. 【実績】昨日までのコストを日別に計算
    for (let d = new Date(startDate); d < todayForCompare; d.setDate(d.getDate() + 1)) {
      if (d > endDate) break; // 念のため月末チェック
      
      if (typeof NewDailyCostCalculator !== 'undefined') {
        // ★ ルーターへ委譲（内部で2026年4月を境に新旧処理を自動で振り分けます）
        NewDailyCostCalculator.routeAndCalculateDailyCost(d, allDailyData, settings, monthlyCache);
      } else if (typeof _calculateDailyPersonnelCosts === 'function') {
        _calculateDailyPersonnelCosts(d, allDailyData, settings); // 移行前フォールバック
      }
    }
    Logger.log(`昨日までの実績コスト計算が完了しました。`);

    // 6. 【予測】本日以降のコストを日別に計算
    const NEW_SYSTEM_START_DATE = new Date('2026-04-01T00:00:00+09:00');
    
    if (startDate >= NEW_SYSTEM_START_DATE && typeof NewDailyCostCalculator !== 'undefined') {
      // ★ 新仕様 (2026/04以降): 確定シフトに予定が入力済みのため、実績と全く同じ高精度エンジンで算出
      for (let d = new Date(todayForCompare); d <= endDate; d.setDate(d.getDate() + 1)) {
        NewDailyCostCalculator.routeAndCalculateDailyCost(d, allDailyData, settings, monthlyCache);
      }
    } else {
      // ★ 旧仕様 (2026/03以前): 過去の予測推計アルゴリズムを使用
      if (typeof _calculateFutureCostsForDailyTool === 'function') {
        const futureCosts = _calculateFutureCostsForDailyTool(startDate, endDate, settings);
        for (const dateKey in futureCosts) {
          if (!allDailyData[dateKey]) allDailyData[dateKey] = {};
          for (const clinicKey in futureCosts[dateKey]) {
            allDailyData[dateKey][clinicKey] = futureCosts[dateKey][clinicKey];
          }
        }
      }
    }
    Logger.log(`本日以降の予測コスト計算が完了しました。`);
    
    // 7. 計算結果をシートに書き出す
    const outputSheet = SPREADSHEET.getSheetByName("拠点別コスト");
    if (!outputSheet) {
      throw new Error("書き込み先シート「拠点別コスト」が見つかりません。");
    }

    _writeResultsToSheet(allDailyData, startDate, endDate, outputSheet, settings, settings.holidays, totalDaysInMonth);
    
    Logger.log(`処理完了: 「拠点別コスト」シートへの書き込みが正常に終了しました。`);

  } catch (e) {
    Logger.log(`エラーが発生したため処理を中断しました: ${e.message}\n${e.stack || ''}`);
  }
}