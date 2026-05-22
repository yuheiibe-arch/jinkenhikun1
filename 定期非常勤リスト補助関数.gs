// =================================================================
// ■ データ取得・整形関数 (ここから不足していた補助関数です)
// =================================================================
function getSourceInfo(urlListData, nameToFind) {
  const nameToFindProcessed = nameToFind.replace(/\s/g, '');
  for (const row of urlListData) {
    if (row[0] && row[0].toString().replace(/\s/g, '') === nameToFindProcessed) {
      return { url: row[1], sheetName: row[2] };
    }
  }
  Logger.log(`警告: URLリストに「${nameToFind}」が見つかりません。`);
  return null;
}

function getAnnualIncomeDataMap(urlListData) {
  const sourceInfo = getSourceInfo(urlListData, "常勤年収");
  if (!sourceInfo) return new Map();
  const sheet = SpreadsheetApp.openByUrl(sourceInfo.url).getSheetByName(sourceInfo.sheetName);
  const header = sheet.getRange(3, 1, 1, sheet.getLastColumn()).getValues()[0];
  const nameCol = header.indexOf("名前") + 1;
  const weeklyHoursCol = header.indexOf("週合計労働時間") + 1;
  const annualHoursCol = header.indexOf("年間総労働時間") + 1;
  const avgWageCol = header.indexOf("平均時給") + 1;
  if (nameCol === 0 || weeklyHoursCol === 0 || annualHoursCol === 0 || avgWageCol === 0) {
    throw new Error("「常勤年収」シートに必要なヘッダーが見つかりません。");
  }
  const data = sheet.getRange(4, 1, sheet.getLastRow() - 3, sheet.getLastColumn()).getValues();
  const dataMap = new Map();
  for (const row of data) {
    const originalName = row[nameCol - 1];
    if (!originalName) continue;
    const cleaned = cleanName(originalName);
    dataMap.set(cleaned, {
      originalName: originalName,
      weeklyHours: row[weeklyHoursCol - 1],
      annualHours: row[annualHoursCol - 1],
      averageHourlyWage: row[avgWageCol - 1]
    });
  }
  return dataMap;
}

function getPhysicianListData(urlListData, year) {
  const sourceInfo = getSourceInfo(urlListData, `定期勤務医師リスト_${year}年度`);
  if (!sourceInfo) return null;
  const sheet = SpreadsheetApp.openByUrl(sourceInfo.url).getSheetByName(sourceInfo.sheetName);
  return sheet.getDataRange().getValues();
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

function createNameCorrectionMap() {
  const correctionMap = new Map([
    ["野呂惠子", "野呂 恵子"], ["藤井泰志", "藤井 泰志①"], ["西宮藤彦", "西宮 藤彦パターン①"],
    ["中島圭代", "中島 圭代（評価変更B→A）"], ["長﨑翔", "長﨑翔（評価変更）"], ["森下あおい", "森下 あおい（評価変更B→A）"],
    ["石井淳子", "石井 淳子　②"], ["渡邊康博", "渡邊 康博（評価変更B→A）"], ["太田みゆき", "太田 みゆき　8月から"],
    ["寺原朋裕", "寺原 朋裕5月から"], ["太田充彦", "太田 充彦（現状維持）"], ["長島由佳", "長島 由佳（評価変更B→A）"],
    ["藤井聡子", "藤井 聡子（祝日なし）"], ["藤原隆弘", "W藤原 隆弘（土曜午前あり）"], ["山本敬一", "山本 敬一①"],
    ["今井健太", "今井 健太（管理医師）"], ["川﨑達人", "川﨑 達人　7月以降"], ["佐野正太郎", "佐野 正太郎（管理医師）8月から"],
    ["中河秀憲", "中河 秀憲（管理医師）"], ["野村莉紗", "野村 莉紗8月から"], ["秋林雅也", "秋林 雅也（週4日）"],
    ["渡邉泰二郎", "渡邉 泰二郎"], ["渡邉貴明", "渡邉 貴明"],
  ]);
  return correctionMap;
}

function cleanName(name) {
  if (typeof name !== 'string' || !name) return '';
  let processedName = name.toString().trim();
  processedName = processedName.replace(/（[^）]*）|\([^)]*\)|[①②③④⑤⑥⑦⑧⑨⑩]/g, '');
  processedName = processedName.replace(/\d{1,2}月から|\d{1,2}月以降/g, '');
  processedName = processedName.replace(/パターン\S*/g, '');
  processedName = processedName.replace(/[！-～]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));
  processedName = processedName.replace(/[\u30a1-\u30f6]/g, s => String.fromCharCode(s.charCodeAt(0) - 0x60));
  processedName = processedName.replace(/^[●Ww\s]+/, '').replace(/\s/g, '');
  return processedName.trim();
}