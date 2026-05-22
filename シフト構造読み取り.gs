/**
 * =================================================================
 * 【検証用】正しいロジックでシフトを読み取り、勤務内訳を集計する (最終版)
 * =================================================================
 */
function debugReadSpecificClinicShift_Final() {
  const SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
  
  // --- ▼設定項目▼ ---
  const TARGET_OFFICIAL_NAME = '国立';
  const TARGET_MONTH = 8; // 9月
  // --- ▲設定項目▲ ---

  const aliasMap = { '流山おおたかの森': '流山', '八千代緑が丘': '八千代', '新鎌ケ谷': '新鎌ヶ谷' };

  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(`「${TARGET_OFFICIAL_NAME}」の${TARGET_MONTH + 1}月分シフト読み取りを開始...`);
    Logger.log(`--- 「${TARGET_OFFICIAL_NAME}」 (${TARGET_MONTH + 1}月) の読み取り検証 (最終版) ---`);

    const urlListSheet = SPREADSHEET.getSheetByName('URLリスト');
    const urlListData = urlListSheet.getDataRange().getValues();
    const kantoEntry = urlListData.find(row => row[0] === '2025関東シフト表');
    const kansaiEntry = urlListData.find(row => row[0] === '2025関西シフト表');
    const kantoSs = SpreadsheetApp.openByUrl(kantoEntry[1]);
    const kansaiSs = SpreadsheetApp.openByUrl(kansaiEntry[1]);
    const allSheets = [...kantoSs.getSheets(), ...kansaiSs.getSheets()];
    
    const searchTerm = aliasMap[TARGET_OFFICIAL_NAME] || TARGET_OFFICIAL_NAME;
    const targetSheet = allSheets.find(sheet => sheet.getName().includes(searchTerm));
    if (!targetSheet) throw new Error(`シートが見つかりません: ${TARGET_OFFICIAL_NAME}`);
    
    const actualSheetName = targetSheet.getName();
    Logger.log(`シート「${actualSheetName}」を特定。`);

    const data = targetSheet.getDataRange().getValues();
    const dateRow = data[24];
    let monthStartCol = -1, targetYear = null;
    for (let c = 0; c < dateRow.length; c++) {
      const cell = dateRow[c];
      if (cell instanceof Date && cell.getMonth() === TARGET_MONTH && cell.getDate() === 1) {
        monthStartCol = c;
        targetYear = cell.getFullYear();
        break;
      }
    }
    if (monthStartCol === -1) throw new Error(`${TARGET_MONTH + 1}月のデータが見つかりません。`);
    Logger.log(`${targetYear}年${TARGET_MONTH + 1}月のデータ（${monthStartCol + 1}列目〜）を読み取ります。`);

    const shiftDetails = {}; 
    const hourMapping = [9, 10, 11, 12, null, null, 15, 16, 17, 18, 19, 20];

    for (let r = 28; r < data.length; r += 2) {
      const ruleRow1 = data[r];
      const ruleRow2 = data[r + 1];
      const dayOfWeekStr = ruleRow1[monthStartCol - 3];
      const weekNumStr = ruleRow1[monthStartCol - 2];
      if (!dayOfWeekStr || weekNumStr === '' || weekNumStr === null) continue;

      const ruleText = (String(weekNumStr).includes('毎週') ? '毎週' : `第${weekNumStr}`) + ` ${dayOfWeekStr}`;
      const dates = _calculateDatesFromRule(targetYear, TARGET_MONTH, weekNumStr, dayOfWeekStr);

      // ▼▼▼【ロジック修正】▼▼▼
      // ルールに該当する日付の「数」だけ、時間と担当者を読み込む
      if (dates.length > 0) {
        for (let c_offset = 0; c_offset < hourMapping.length; c_offset++) {
          if (hourMapping[c_offset] === null) continue;
          
          // 担当者名は常に「月の開始列」から読み取る
          const persons = [
            (ruleRow1[monthStartCol + c_offset] || '').toString().trim(),
            (ruleRow2 && ruleRow2[monthStartCol + c_offset] ? ruleRow2[monthStartCol + c_offset] : '').toString().trim()
          ];
          
          persons.forEach(person => {
            if (_isValidName(person)) {
              if (!shiftDetails[person]) shiftDetails[person] = {};
              if (!shiftDetails[person][ruleText]) shiftDetails[person][ruleText] = 0;
              // ルールが適用される日数分だけ時間を加算
              shiftDetails[person][ruleText] += dates.length;
            }
          });
        }
      }
    }

    Logger.log(`\n--- 「${actualSheetName}」 ${targetYear}年${TARGET_MONTH + 1}月分 集計結果 ---`);
    const sortedNames = Object.keys(shiftDetails).sort();
    if (sortedNames.length === 0) {
      Logger.log("該当月のシフトデータが見つかりませんでした。");
    } else {
      sortedNames.forEach(name => {
        Logger.log(`■ ${name}`);
        let totalHours = 0;
        const rules = shiftDetails[name];
        for (const rule in rules) {
          const hours = rules[rule];
          totalHours += hours;
          Logger.log(`    - ${rule}: ${hours} 時間`);
        }
        Logger.log(`  【合計: ${totalHours} 時間】\n`);
      });
    }
    Logger.log("--------------------------------------------------");
    
    SpreadsheetApp.getActiveSpreadsheet().toast('読み取りが完了しました。ログを確認してください。');

  } catch (e) {
    Logger.log(`エラーが発生しました: ${e.message}`);
    Logger.log(`Stack: ${e.stack}`);
    SpreadsheetApp.getActiveSpreadsheet().toast(`エラー: ${e.message}`);
  }
}

function _isValidName(name) {
  if (!name || name.length <= 1) return false;
  if (name.includes('募集')) return true;
  if (/\d/.test(name)) return false;
  if (/診目$/.test(name)) return false;
  if (/^(月|火|水|木|金|土|日)$/.test(name)) return false;
  return true;
}

function _calculateDatesFromRule(year, month, weekNumStr, dayOfWeekStr) {
  const weekRule = String(weekNumStr);
  const dayMap = { "日": 0, "月": 1, "火": 2, "水": 3, "木": 4, "金": 5, "土": 6 };
  const targetDayOfWeek = dayMap[dayOfWeekStr.charAt(0)];
  if (targetDayOfWeek === undefined) return [];

  const dates = [];
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const weeks = weekRule.includes('毎週') ? null : weekRule.match(/\d/g).map(Number);

  for (let day = 1; day <= daysInMonth; day++) {
    const d = new Date(year, month, day);
    if (d.getDay() === targetDayOfWeek) {
      if (weeks) {
        const nthDay = Math.floor((d.getDate() - 1) / 7) + 1;
        if (weeks.includes(nthDay)) {
          dates.push(d);
        }
      } else {
        dates.push(d);
      }
    }
  }
  return dates;
}