/**
 * =================================================================
 * 人件費計算スクリプト メイン処理
 * =================================================================
 */
function masterOrchestrator() {
  const SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
  const URL_LIST_SHEET_NAME = "URLリスト";
  const HOLIDAY_SHEET_NAME = "祝日";
  const OUTPUT_SHEET_NAME = "常勤マスタ";

  try {
    Logger.log("--- 処理を開始します ---");
    SpreadsheetApp.getActive().toast("処理を開始しました。完了まで数分かかる場合があります。");

    const urlListSheet = SPREADSHEET.getSheetByName(URL_LIST_SHEET_NAME);
    if (!urlListSheet) throw new Error(`シート「${URL_LIST_SHEET_NAME}」が見つかりません。`);
    
    const urlListData = urlListSheet.getDataRange().getValues();
    const holidayList = getHolidaySet(SPREADSHEET.getSheetByName(HOLIDAY_SHEET_NAME));
    const annualIncomeDataMap = getAnnualIncomeDataMap(urlListData);
    const nameCorrectionMap = createNameCorrectionMap();

    Logger.log(`祝日リスト、常勤年収データ、氏名修正候補を読み込みました。`);

    let finalOutputData = [];
    const targetYears = [2025, 2024];

    for (const year of targetYears) {
      Logger.log(`--- ${year}年度の処理を開始 ---`);
      const physicianListData = getPhysicianListData(urlListData, year);
      
      if (physicianListData) {
        const processedData = processPhysicianList({
          year: year,
          physicianListData: physicianListData,
          annualIncomeDataMap: annualIncomeDataMap,
          holidayList: holidayList,
          nameCorrectionMap: nameCorrectionMap
        });
        finalOutputData = finalOutputData.concat(processedData);
        Logger.log(`${year}年度の処理が完了。${processedData.length}件の常勤医師データを生成しました。`);
      }
    }

    if (finalOutputData.length > 0) {
      const outputSheet = SPREADSHEET.getSheetByName(OUTPUT_SHEET_NAME);
      writeToMasterSheet(outputSheet, finalOutputData);
      
      Logger.log(`--- 全ての処理が完了しました ---`);
      SpreadsheetApp.getActive().toast(`処理が完了しました。常勤マスタに${finalOutputData.length}件のデータを出力しました。`);
    } else {
      Logger.log("処理対象のデータが見つかりませんでした。");
      SpreadsheetApp.getActive().toast("処理対象のデータが見つかりませんでした。");
    }

  } catch (e) {
    Logger.log(`エラーが発生しました: ${e.message} \nStack: ${e.stack}`);
    SpreadsheetApp.getUi().alert(`処理中にエラーが発生しました。\n\n詳細:\n${e.message}`);
  }
}

/**
 * AC列（月別サマリー）とAD列（差異チェック）のデータを生成する
 * @param {object} calcResult - calculateMonthlyHoursFromContractからの戻り値。monthlyWorkDays配列を含む想定。
 * @param {object} baseInfo - annualIncomeDataMapから取得した基本情報
 * @param {number} year - 対象年度
 * @returns {{summaryText: string, verificationMessage: string}}
 */
function createAdditionalColumns(calcResult, baseInfo, year) {
  // --- AC列: 月別サマリー ---
  const summaryLines = [];
  // 前提条件：calcResultが月別の日数配列を持っているか確認
  if (calcResult.monthlyWorkDays && calcResult.monthlyWorkDays.length === 12) {
    for (let i = 0; i < 12; i++) {
      const month = (i + 3) % 12 + 1;
      const targetYear = (month < 4) ? year + 1 : year;
      const days = calcResult.monthlyWorkDays[i];
      const hours = calcResult.monthlyHours[i];
      if (hours > 0) { // 勤務時間がある月のみ表示
        const monthStr = String(month).padStart(2, '0');
        summaryLines.push(`${targetYear}/${monthStr}:${days}d(${hours.toFixed(2)}h)`);
      }
    }
  }
  const summaryText = summaryLines.join("\n");

  // --- AD列: 年間時間差異チェック ---
  const calculatedTotalHours = calcResult.monthlyHours.reduce((sum, h) => sum + h, 0);
  const masterAnnualHours = baseInfo.annualHours || 0;
  let verificationMessage = "";
  
  // 浮動小数点誤差を考慮して比較
  const difference = Math.round((calculatedTotalHours - masterAnnualHours) * 100) / 100;

  if (difference > 0) {
    verificationMessage = `勤務合計が ${difference}h 多い`;
  } else if (difference < 0) {
    verificationMessage = `勤務合計が ${Math.abs(difference)}h 少ない`;
  } else {
    verificationMessage = "一致";
  }
  
  return { summaryText, verificationMessage };
}


function processPhysicianList(params) {
  const { year, physicianListData, annualIncomeDataMap, holidayList, nameCorrectionMap } = params;
  
  const header = physicianListData[4];
  const dataRows = physicianListData.slice(5);
  const outputData = [];
  
  const colIdx = {
    doctorCategory: header.indexOf("医師区分"),
    licenseNum: header.indexOf("医籍番号"),
    name: header.indexOf("氏名\nスペース\nいれない！！"),
    mainWorkplace: header.indexOf("主務"),
    specialty: header.indexOf("診療科"),
    joinDate: header.indexOf("入職日"),
    contract: header.indexOf(`${year}年度契約内容`),
    holidayWork: header.indexOf("祝日"),
    newYearWork: header.indexOf("年末年始")
  };

  for (const row of dataRows) {
    if (row[colIdx.doctorCategory] !== "常勤" || !row[colIdx.name]) {
      continue;
    }

    const originalPhysicianName = row[colIdx.name];
    const physicianNameToSearch = nameCorrectionMap.get(originalPhysicianName) || originalPhysicianName;
    const cleanedName = cleanName(physicianNameToSearch);
    
    const baseInfo = annualIncomeDataMap.get(cleanedName);

    if (!baseInfo) {
      Logger.log(`警告: ${year}年度の「${originalPhysicianName}」は、常勤年収データに見つかりませんでした。`);
      continue;
    }
    
    const calcResult = calculateMonthlyHoursFromContract(
      row[colIdx.contract], 
      year, 
      row[colIdx.holidayWork] !== "無", 
      holidayList,
      row[colIdx.newYearWork] !== "無"
    );

    const averageHourlyWage = baseInfo.averageHourlyWage ? Math.round(baseInfo.averageHourlyWage) : '';
    const finalCleanedName = cleanName(baseInfo.originalName);
    
    const { summaryText, verificationMessage } = createAdditionalColumns(calcResult, baseInfo, year);

    outputData.push([
      null,
      row[colIdx.licenseNum],
      finalCleanedName,
      row[colIdx.mainWorkplace],
      row[colIdx.contract],
      row[colIdx.joinDate],
      row[colIdx.specialty],
      baseInfo.weeklyHours,
      baseInfo.annualHours,
      row[colIdx.holidayWork],
      row[colIdx.newYearWork],
      year,
      ...calcResult.monthlyHours,
      calcResult.totalWorkDays,
      averageHourlyWage,
      '',
      '',
      summaryText,
      verificationMessage
    ]);

    // ---【特例処理】---
    const licenseNum = row[colIdx.licenseNum].toString();

    // 野呂恵子さんの特例処理 (医籍番号: 330819)
    if (licenseNum === "330819") {
      const originalContract = row[colIdx.contract];
      const specialContract = originalContract.replace(/【代官山】第[24]土曜日\s*9:00～21:00/g, "【代官山】毎週土曜日　9:00～21:00");
      
      const specialCalcResult = calculateMonthlyHoursFromContract(
        specialContract, year, row[colIdx.holidayWork] !== "無", holidayList, row[colIdx.newYearWork] !== "無"
      );
      
      const { summaryText: specialSummary, verificationMessage: specialVerification } = createAdditionalColumns(specialCalcResult, baseInfo, year);

      const specialRow = [...outputData[outputData.length - 1]]; 
      specialRow[2] = `（特例）${finalCleanedName}`;
      specialRow[4] = specialContract;
      for (let i = 0; i < 12; i++) {
        specialRow[12 + i] = specialCalcResult.monthlyHours[i];
      }
      specialRow[24] = specialCalcResult.totalWorkDays;
      specialRow[28] = specialSummary;
      specialRow[29] = specialVerification;
      outputData.push(specialRow);
    }

    // 中村千穂さんの特例処理 (医籍番号: 496210)
    if (licenseNum === "496210") {
      const originalContract = row[colIdx.contract];
      const specialContract = originalContract.replace(/【流山】第[135]水曜日\s*9:00～13:00/g, "【流山】毎週水曜日　9:00～13:00");

      const specialCalcResult = calculateMonthlyHoursFromContract(
        specialContract, year, row[colIdx.holidayWork] !== "無", holidayList, row[colIdx.newYearWork] !== "無"
      );
      
      const { summaryText: specialSummary, verificationMessage: specialVerification } = createAdditionalColumns(specialCalcResult, baseInfo, year);

      const specialRow = [...outputData[outputData.length - 1]];
      specialRow[2] = `（特例）${finalCleanedName}`;
      specialRow[4] = specialContract;
      for (let i = 0; i < 12; i++) {
        specialRow[12 + i] = specialCalcResult.monthlyHours[i];
      }
      specialRow[24] = specialCalcResult.totalWorkDays;
      specialRow[28] = specialSummary;
      specialRow[29] = specialVerification;
      outputData.push(specialRow);
    }

    // 寺原朋裕さんの特例処理 (医籍番号: 430324)
    if (licenseNum === "430324") {
      const originalContract = row[colIdx.contract];
      const specialContract = originalContract
        .replace(/【村上】毎週水曜日：9:00～13:00\s*└9:00～13:00は拠点ご勤務/g, "【複数】毎週水曜日：9:00～19:00")
        .replace(/【MQC】毎週水曜日：15:00～21:00\s*└15:00～21:00はMQC業務とTS会議参加/g, "");

      const specialCalcResult = calculateMonthlyHoursFromContract(
        specialContract, year, row[colIdx.holidayWork] !== "無", holidayList, row[colIdx.newYearWork] !== "無"
      );
      
      const { summaryText: specialSummary, verificationMessage: specialVerification } = createAdditionalColumns(specialCalcResult, baseInfo, year);
      
      const specialRow = [...outputData[outputData.length - 1]]; 
      specialRow[2] = `（特例）${finalCleanedName}`;
      specialRow[4] = specialContract.trim();
      for (let i = 0; i < 12; i++) {
        specialRow[12 + i] = specialCalcResult.monthlyHours[i];
      }
      specialRow[24] = specialCalcResult.totalWorkDays;
      specialRow[28] = specialSummary;
      specialRow[29] = specialVerification;
      outputData.push(specialRow);
    }
  }
  
  return outputData.map((row, index) => {
    row[0] = index + 1;
    return row;
  });
}


function writeToMasterSheet(sheet, data) {
  const header = [
    "番号", "医籍番号", "医師名", "主勤務先", "契約内容", "入職日", "専門", 
    "週間労働時間", "年間労働時間", "土日祝稼働", "年末年始稼働", "年度",
    "4月", "5月", "6月", "7月", "8月", "9月", 
    "10月", "11月", "12月", "1月", "2月", "3月",
    "延べ出勤日数", "平均時給", "交通費", "そのほか支出",
    "月別勤務サマリー", "年間時間差異チェック"
  ];
  
  sheet.clear();
  sheet.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight("bold");
  
  if (data && data.length > 0) {
    sheet.getRange(2, 29, data.length, 1).setWrap(true); // AC列の折り返しを有効に
    sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  }
  
  Logger.log(`${sheet.getName()}シートへの書き込みが完了しました。`);
}