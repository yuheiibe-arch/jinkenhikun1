/**
 * ====================================================================
 * 【ファイル: 04_NewDataFetcher.gs】(修正版)
 * 2026年度以降の新仕様シートからデータを取得・整形するための専用モジュール
 * ====================================================================
 */

const NewDataFetcher = (function() {
  
  function _getEntryFromUrlList(targetName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const urlSheet = ss.getSheetByName("URLリスト");
    if (!urlSheet) throw new Error("「URLリスト」シートが見つかりません。");
    
    const data = urlSheet.getDataRange().getValues();
    const entry = data.find(row => row[0] && row[0].toString().trim() === targetName);
    
    if (!entry || !entry[1]) {
      throw new Error(`URLリストに「${targetName}」が登録されていない、またはURLが空です。`);
    }
    
    return {
      url: entry[1].toString().trim(),
      sheetName: entry[2] ? entry[2].toString().trim() : null
    };
  }

  // ★ 修正: 日本時間 (JST) で強制的に "HH:mm" に変換して時間ズレを防ぐ
  function _convertSerialTimeToString(timeValue) {
    if (!timeValue) return "";
    if (timeValue instanceof Date) {
      return Utilities.formatDate(timeValue, "JST", "HH:mm");
    }
    return String(timeValue).trim();
  }

  function fetchShiftData(startDate, endDate) {
    const entry = _getEntryFromUrlList("確定シフト");
    const ss = SpreadsheetApp.openByUrl(entry.url);
    const sheet = ss.getSheetByName("確定シフト");
    if (!sheet) throw new Error(`確定シフトのスプレッドシートに「確定シフト」シートが存在しません。`);

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return []; 

    const header = data[0];
    const colIdx = {
      name: header.indexOf("名前"),
      medId: header.indexOf("医籍番号"),
      clinic: header.indexOf("クリニック名"),
      dept: header.indexOf("診療科"),
      date: header.indexOf("勤務日"),
      startTime: header.indexOf("勤務開始時間"),
      endTime: header.indexOf("勤務終了時間"),
      comment1: header.indexOf("スタッフコメント1"),
      wageTotal: header.indexOf("時給合計")
    };

    if (colIdx.date === -1 || colIdx.startTime === -1) {
      throw new Error("確定シフトのヘッダーに必要なカラムが見つかりません。");
    }

    const results = [];
    
    // ★ 修正: 比較も日本時間 (JST) の文字列で行う
    const startStr = Utilities.formatDate(startDate, "JST", "yyyy/MM/dd");
    const endStr = Utilities.formatDate(endDate, "JST", "yyyy/MM/dd");

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const workDate = row[colIdx.date];

      if (!(workDate instanceof Date)) continue;

      const workDateStr = Utilities.formatDate(workDate, "JST", "yyyy/MM/dd");
      
      // 文字列比較で確実に期間内のシフトを抽出する
      if (workDateStr >= startStr && workDateStr <= endStr) {
        results.push({
          rowNumber: i + 1,
          doctorName: row[colIdx.name],
          medId: row[colIdx.medId],
          clinicName: row[colIdx.clinic],
          department: row[colIdx.dept],
          workDate: workDate,
          workDateStr: workDateStr, // ★ これをルーター側での照合に使う
          startTimeStr: _convertSerialTimeToString(row[colIdx.startTime]),
          endTimeStr: _convertSerialTimeToString(row[colIdx.endTime]),
          comment1: row[colIdx.comment1],
          wageTotal: row[colIdx.wageTotal]
        });
      }
    }
    return results;
  }

  function fetchFullTimeMaster2026() {
    const entry = _getEntryFromUrlList("常勤2026年度");
    const ss = SpreadsheetApp.openByUrl(entry.url);
    const sheet = entry.sheetName ? ss.getSheetByName(entry.sheetName) : ss.getSheets()[0];
    if (!sheet) throw new Error(`常勤マスタが見つかりません。`);
    return sheet.getDataRange().getValues();
  }

  function fetchPartTimeMaster2026() {
    const entry = _getEntryFromUrlList("定期勤務医師リスト_2026年度");
    const ss = SpreadsheetApp.openByUrl(entry.url);
    const sheet = entry.sheetName ? ss.getSheetByName(entry.sheetName) : ss.getSheets()[0];
    if (!sheet) throw new Error(`定期非常勤マスタが見つかりません。`);
    return sheet.getDataRange().getValues();
  }

  return {
    fetchShiftData: fetchShiftData,
    fetchFullTimeMaster2026: fetchFullTimeMaster2026,
    fetchPartTimeMaster2026: fetchPartTimeMaster2026
  };
})();