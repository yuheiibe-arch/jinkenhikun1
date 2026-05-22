/**
 * ====================================================================
 * 【検証用】2026年度 新マスタ＆クリニックID＆千葉NTシフト デバッグ
 * ====================================================================
 */
function debugVerifyNewMasterAndID() {
  Logger.log("=== 🚀 新マスタ＆ID照合＆シフト抽出の検証を開始します ===");

  // ============================================================
  // ① 定期非常勤マスタの「契約時給」と「特別時給の内訳」の確認
  // ============================================================
  Logger.log("\n【1】定期非常勤マスタ (2026年度) の読み取り検証");
  try {
    const ptUrl = "https://docs.google.com/spreadsheets/d/1aEjphEv_63SeWQmwiOy9sx7IrMfawU01sHbKd_Ki4iA/edit";
    const ptSheet = SpreadsheetApp.openByUrl(ptUrl).getSheetByName("定期非常勤2026年度");
    const ptData = ptSheet.getDataRange().getValues();
    const ptHeaders = ptData[0];
    
    // 動的インデックス取得
    const hMap = {};
    ptHeaders.forEach((h, idx) => { if (h) hMap[String(h).replace(/\n/g, '').trim()] = idx; });
    
    Logger.log(`✅ 動的ヘッダー取得成功: [契約時給=${hMap['契約時給']}, 特別時給の内訳=${hMap['特別時給の内訳']}]`);

    let foundNaba = false;
    for (let i = 1; i < ptData.length; i++) {
      const name = ptData[i][hMap['医師名']];
      if (name && name.includes("那波")) {
        foundNaba = true;
        Logger.log(`  👨‍⚕️ 医師名: ${name}`);
        Logger.log(`    -> 契約時給: 「${ptData[i][hMap['契約時給']]}」`);
        const specialWageText = ptData[i][hMap['特別時給の内訳']];
        Logger.log(`    -> 特別時給の内訳:\n${specialWageText ? specialWageText : "(空欄)"}`);
      }
    }
    if (!foundNaba) Logger.log("  ⚠️ 那波先生のデータが見つかりませんでした。");
  } catch(e) {
    Logger.log(`❌ 定期非常勤マスタの検証エラー: ${e.message}`);
  }

  // ============================================================
  // ② 時給マスタの「動的シート検索」と「クリニックID」の確認
  // ============================================================
  Logger.log("\n【2】時給マスタ (2026上期/下期) のID紐付け検証");
  try {
    const wageUrl = "https://docs.google.com/spreadsheets/d/14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs/edit";
    const wageSs = SpreadsheetApp.openByUrl(wageUrl);
    const allSheets = wageSs.getSheets();
    
    const targetSheets = allSheets.filter(s => s.getName().endsWith('時給'));
    Logger.log(`✅ 動的シート検索成功: 該当シート [${targetSheets.map(s => s.getName()).join(', ')}]`);

    targetSheets.forEach(sheet => {
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const hMap = {};
      headers.forEach((h, idx) => { if (h) hMap[String(h).trim()] = idx; });
      
      Logger.log(`  📄 シート「${sheet.getName()}」 (クリニックID列: ${hMap['クリニックID'] !== undefined ? 'あり' : 'なし'})`);
      
      // サンプルとして3件だけIDと拠点名を出力
      let count = 0;
      for (let i = 1; i < data.length && count < 3; i++) {
        const cId = hMap['クリニックID'] !== undefined ? data[i][hMap['クリニックID']] : "";
        const cName = hMap['拠点名'] !== undefined ? data[i][hMap['拠点名']] : "";
        const cDept = hMap['科目'] !== undefined ? data[i][hMap['科目']] : "";
        
        if (cName || cId) {
          Logger.log(`    -> ID: [${cId}] | 拠点: ${cName} | 科目: ${cDept} | 平日午前: ${data[i][hMap['平日_午前']]}円`);
          count++;
        }
      }
    });
  } catch(e) {
    Logger.log(`❌ 時給マスタの検証エラー: ${e.message}`);
  }

  // ============================================================
  // ③ 千葉ニュータウン (千葉NT) のシフト抽出確認
  // ============================================================
  Logger.log("\n【3】千葉ニュータウンのシフト抽出 (表記ブレ吸収) 検証");
  try {
    const TARGET_DATE = new Date('2026-04-13T00:00:00+09:00');
    const settings = _loadInitialData(SpreadsheetApp.getActiveSpreadsheet(), TARGET_DATE);
    
    // キャッシュから4月13日のシフトを取得
    const cache = NewDailyCostCalculator.buildMonthlyCache(
      new Date('2026-04-01T00:00:00+09:00'), 
      new Date('2026-04-30T00:00:00+09:00')
    );
    const dateStr = Utilities.formatDate(TARGET_DATE, 'JST', 'yyyy/MM/dd');
    const dailyShifts = cache.shifts.filter(s => s.workDateStr === dateStr);

    Logger.log(`✅ 4/13の全シフト数: ${dailyShifts.length}件`);

    // ★修正ポイント: シフト側の拠点名も正規化してから比較する
    const chibaNtShifts = dailyShifts.filter(s => {
       const rawName = String(s.clinicName);
       const normName = gt_normalizeClinicName(rawName, settings.clinicAliasMap);
       return normName === '千葉ニュータウン' || rawName.includes('千葉NT') || rawName.includes('千葉ニュータウン');
    });

    Logger.log(`✅ 千葉ニュータウン関連のシフト抽出: ${chibaNtShifts.length}件`);
    
    chibaNtShifts.forEach(s => {
      const normName = gt_normalizePersonName(s.doctorName, settings.nameAliasMap);
      Logger.log(`  👨‍⚕️ 医師: ${s.doctorName} (正規化: ${normName})`);
      Logger.log(`    -> 生の拠点名: 「${s.clinicName}」`);
      Logger.log(`    -> 時間: ${s.startTimeStr} - ${s.endTimeStr}`);
      
      // 常勤マスタでの時給登録状況を確認
      const ftWage = settings.ftHourlyWageMap.get(normName);
      if (ftWage) {
        Logger.log(`    -> 💡 常勤マスタ時給登録あり: ${ftWage}円`);
      } else {
        Logger.log(`    -> 💡 常勤マスタ時給登録なし`);
      }
    });

  } catch(e) {
    Logger.log(`❌ シフト抽出の検証エラー: ${e.message}`);
  }

  Logger.log("\n=== 🚀 デバッグ完了 ===");
}