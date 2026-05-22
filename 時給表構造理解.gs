/**
 * =================================================================
 * 【調査用】指定された条件の時給をマスターから正しく読み取れるか調査する
 * =================================================================
 */
function debugGetHourlyWage() {
  // --- ▼調査したい条件を指定▼ ---
  const TARGET_CLINIC = '国立';
  const TARGET_DEPARTMENT = '小児科'; // ※診療科が異なる場合は修正してください
  const TARGET_DATE = new Date('2025-09-14T15:00:00'); // 日曜日であればいつでもOK
  const TARGET_SHIFT_SEGMENT = { name: 'pm' }; // 15:00-18:00は 'pm' (午後)に該当
  // --- ▲調査したい条件を指定▲ ---

  try {
    Logger.log(`--- 時給読み取り調査を開始 ---`);
    Logger.log(`条件: ${TARGET_CLINIC} / ${TARGET_DEPARTMENT} / ${Utilities.formatDate(TARGET_DATE, "JST", "yyyy/MM/dd(E)")} / ${TARGET_SHIFT_SEGMENT.name.toUpperCase()}シフト`);

    // 1. 判定に必要なマスターデータと祝日リストを読み込む
    const payRateMaster = getPayRateMaster();
    const holidays = getHolidays();
    Logger.log("✅ ステップ1: 時給マスターと祝日リストの読み込み完了");

    // 2. 指定された日付の曜日タイプを判定する
    const dayType = getDayType(TARGET_DATE, holidays);
    if (!dayType) throw new Error("曜日タイプを判定できませんでした。");
    Logger.log(`✅ ステップ2: 曜日タイプは「${dayType}」と判定`);

    // 3. 時給マスターから、該当クリニック・診療科の時給テーブルを取得
    const masterKey = `${TARGET_CLINIC}||${TARGET_DEPARTMENT}`;
    const clinicPayRates = payRateMaster[masterKey];
    if (!clinicPayRates) {
      // 共通設定も試す
      const commonKey = `${TARGET_CLINIC}||共通`;
      const commonPayRates = payRateMaster[commonKey];
      if(!commonPayRates) {
         throw new Error(`時給マスターに「${TARGET_CLINIC}」の「${TARGET_DEPARTMENT}」または「共通」の設定が見つかりません。`);
      }
      Logger.log(`✅ ステップ3: 「${TARGET_CLINIC}」の「共通」の時給テーブルを取得`);
      clinicPayRates = commonPayRates;
    } else {
       Logger.log(`✅ ステップ3: 「${TARGET_CLINIC}」の「${TARGET_DEPARTMENT}」の時給テーブルを取得`);
    }

    // 4. 条件に合致する正しい時給を引き当てる
    const correctPay = getCorrectPay(TARGET_SHIFT_SEGMENT, dayType, clinicPayRates);
    Logger.log(`✅ ステップ4: 時給の引き当て処理を実行`);

    if (correctPay !== null && correctPay !== '') {
      Logger.log(`\n🎉【成功】時給を特定しました！`);
      Logger.log(`  ▶ 時給は ${correctPay} 円です。`);
    } else {
      Logger.log(`\n⚠️【警告】時給を特定できませんでした。`);
      Logger.log(`時給マスターの「${dayType}」の「${TARGET_SHIFT_SEGMENT.name}」欄が空、または設定が存在しない可能性があります。`);
    }

  } catch (e) {
    Logger.log(`\n❌【エラー】処理中に問題が発生しました。`);
    Logger.log(e.message);
  }
}


// ===================================================================
// ===== 以下、ご提示いただいたヘルパー関数群（変更なし）=====
// ===================================================================

function getCorrectPay(segment, dayType, masterRow) {
  if (!segment || !masterRow) return null;
  const type = dayType.toLowerCase();
  const zone = segment.name;
  if (masterRow[type] && masterRow[type][zone]) {
    return masterRow[type][zone];
  }
  return null;
}

function getDayType(date, holidays) {
  const workDate = (date instanceof Date) ? date : new Date(date);
  if (isNaN(workDate.getTime())) return '';
  const dateStr = Utilities.formatDate(workDate, "JST", "yyyy-MM-dd");
  if (holidays.has(dateStr) || workDate.getDay() === 0) return 'holiday';
  if (workDate.getDay() === 6) return 'saturday';
  return 'weekday';
}

function getHolidays() {
  return new Set(['2025-01-01', '2025-01-13', '2025-02-11', '2025-02-23', '2025-02-24', '2025-03-20', '2025-04-29', '2025-05-03', '2025-05-04', '2025-05-05', '2025-05-06', '2025-07-21', '2025-08-11', '2025-09-15', '2025-09-23', '2025-10-13', '2025-11-03', '2025-11-23', '2025-11-24', '2026-01-01', '2026-01-12', '2026-02-11', '2026-02-23', '2026-03-20']);
}

function getPayRateMaster() {
  const MASTER_URL = 'https://docs.google.com/spreadsheets/d/1eqejNaKWSuHVnRwxaGT-RHgOnlsHkcXYQ5J32B8T_XM/edit';
  const MASTER_SHEET_NAME = '2025年度（年間）';
  let masterSs;
  try { masterSs = SpreadsheetApp.openByUrl(MASTER_URL); }  
  catch(e) { throw new Error(`時給マスタのURLにアクセスできませんでした。\n${e.message}`); }
  const masterSheet = masterSs.getSheetByName(MASTER_SHEET_NAME);
  if (!masterSheet) throw new Error(`時給マスタに「${MASTER_SHEET_NAME}」シートが見つかりません。`);
  const values = masterSheet.getRange('B2:O' + masterSheet.getLastRow()).getValues();
  const masterData = {};
  values.forEach(row => {
    const clinic = (row[0] || '').toString().trim();
    const department = (row[1] || '').toString().replace(/\s+/g, '');
    if (!clinic) return;
    const key = `${clinic}||${department}`;
    masterData[key] = {
      weekday: { am: row[3], pm: row[4], night: row[5] },
      saturday: { am: row[7], pm: row[8], night: row[9] },
      holiday: { am: row[11], pm: row[12], night: row[13] }
    };
  });
  return masterData;
}