/**
 * ====================================================================
 * 【ファイル: 06_NewDoctorMasterManager.gs】
 * 2026年度 新・医師マスタ（常勤・定期非常勤）の動的ローダー
 * * 💡【設計仕様】
 * 1. ハードコード禁止: ヘッダー名から動的に列インデックスを取得。列の増減・入替に完全耐性。
 * 2. 退職者の自動除外: 「退職日」を判定し、計算対象日時点で退職済みの医師はメモリに読み込まない。
 * 3. 完全外部参照化: ローカルの古いマスタシートは一切見ず、この外部URLのみを正とする。
 * ====================================================================
 */

const NewDoctorMasterManager = (function() {
  const MASTER_URL = "https://docs.google.com/spreadsheets/d/1aEjphEv_63SeWQmwiOy9sx7IrMfawU01sHbKd_Ki4iA/edit";

  /**
   * ヘッダー行から { "列名": インデックス番号 } のマップを動的に生成する
   */
  function _createHeaderMap(headerRow) {
    const map = {};
    headerRow.forEach((val, idx) => {
      if (val) map[String(val).replace(/\n/g, '').trim()] = idx;
    });
    return map;
  }

  /**
   * 【メインAPI】医師マスタを読み込み、システム全体で使うMap群を生成する
   * @param {Map} nameAliasMap - 氏名の表記揺れ吸収マップ
   * @param {Map} clinicAliasMap - 拠点名の表記揺れ吸収マップ
   * @param {Date} targetDate - 計算対象日（この日以前に退職している医師は除外される）
   */
  function buildDoctorMaps(nameAliasMap, clinicAliasMap, targetDate = new Date()) {
    const ss = SpreadsheetApp.openByUrl(MASTER_URL);

    // 戻り値用オブジェクトの準備
    const ftHourlyWageMap = new Map();
    const ptContractWageMap = new Map();
    const doctorContractMap = new Map();
    const ptDoctorContractMapById = new Map();
    const nameToMedicalIdMap = new Map();
    const partTimeDoctorNameSet = new Set();

    // ====================================================
    // [1] 常勤マスタ (2026年度) の読み込み
    // ====================================================
    const ftSheet = ss.getSheetByName('常勤2026年度');
    if (ftSheet) {
      const data = ftSheet.getDataRange().getValues();
      const hMap = _createHeaderMap(data[0]);

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const rawName = row[hMap['医師名']];
        const rawMedId = row[hMap['医籍番号']];
        const retireDateRaw = row[hMap['退職日']];
        
        if (!rawName) continue;

        // 🌟【退職者除外ロジック】退職日が設定されており、かつ対象日より過去であればスキップ
        if (retireDateRaw instanceof Date && retireDateRaw < targetDate) {
          continue; 
        }

        const normName = gt_normalizePersonName(rawName, nameAliasMap);
        const medId = rawMedId ? String(rawMedId).trim() : null;
        
        // 時給の取得（"不要"などの文字列が入っている場合はデフォルト値10000をセット）
        const wageVal = parseInt(row[hMap['時給']], 10);
        const wage = isNaN(wageVal) ? 10000 : wageVal;

        if (normName) {
          ftHourlyWageMap.set(normName, wage);
          doctorContractMap.set(normName, { isFullTime: true });
          if (medId) nameToMedicalIdMap.set(normName, medId);
        }
      }
    }

    // ====================================================
    // [2] 定期非常勤マスタ (2026年度) の読み込み
    // ====================================================
    const ptSheet = ss.getSheetByName('定期非常勤2026年度');
    if (ptSheet) {
      const data = ptSheet.getDataRange().getValues();
      const hMap = _createHeaderMap(data[0]);

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const rawName = row[hMap['医師名']];
        const rawMedId = row[hMap['医籍番号']];
        const retireDateRaw = row[hMap['退職日']];
        const contractText = row[hMap['勤務備考']];
        
        if (!rawName) continue;

        // 🌟【退職者除外ロジック】
        if (retireDateRaw instanceof Date && retireDateRaw < targetDate) {
          continue; 
        }

        const normName = gt_normalizePersonName(rawName, nameAliasMap);
        const medId = rawMedId ? String(rawMedId).trim() : null;
        
        const wageVal = parseInt(row[hMap['契約時給']], 10);
        const wage = isNaN(wageVal) ? 15000 : wageVal;

        if (normName) {
          partTimeDoctorNameSet.add(normName);
          ptContractWageMap.set(normName, wage);
          if (medId) nameToMedicalIdMap.set(normName, medId);

          // 契約内容（勤務備考）の解析と登録
          if (medId && contractText && typeof parseContractToRules === 'function') {
            const rules = parseContractToRules(contractText);
            
            const startDateRaw = row[hMap['入職日']];
            let startDate = new Date('2000/01/01'); // 安全のためのフォールバック
            if (startDateRaw instanceof Date) {
              startDate = new Date(startDateRaw);
              startDate.setHours(0,0,0,0);
            }

            const contracts = rules.map(rule => ({
              location: gt_normalizeClinicName(rule.workplace, clinicAliasMap),
              dayOfWeek: rule.dayOfWeek
            })).filter(c => c.location && typeof c.dayOfWeek !== 'undefined');

            ptDoctorContractMapById.set(medId, { startDate, contracts, name: normName });
          }
        }
      }
    }

    return {
      ftHourlyWageMap,
      ptContractWageMap,
      doctorContractMap,
      ptDoctorContractMapById,
      partTimeDoctorNameSet,
      nameToMedicalIdMap
    };
  }

  return { buildDoctorMaps };
})();