function runSystemHealthCheck_() {
  const issues = [];
  const warnings = [];
  let jobCount = 0;
  let dancerCount = 0;

  try {
    const masterSs = getMasterSpreadsheet();
    const jobSheet = masterSs.getSheetByName('案件マスタ');
    if (!jobSheet) throw new Error('「案件マスタ」シートがありません');

    const jobs = jobSheet.getDataRange().getValues().slice(1);
    const jobMap = {};
    jobs.forEach(function(row, index) {
      const name = String(row[0] || '').trim();
      if (!name) return;
      jobCount += 1;
      const price = Number(row[2]);
      if (!Number.isFinite(price) || price < 0) {
        issues.push('案件マスタ' + (index + 2) + '行目の単価が不正です: ' + name);
      }
      if (Object.prototype.hasOwnProperty.call(jobMap, name) && jobMap[name] !== price) {
        issues.push('案件マスタに異なる単価の同名案件があります: ' + name);
      }
      jobMap[name] = price;
    });

    const contacts = getOutsourceContacts_();
    dancerCount = contacts.length;
    const seenNames = {};
    contacts.forEach(function(person) {
      if (seenNames[person.stageName]) issues.push('外注連絡票の芸名が重複しています: ' + person.stageName);
      seenNames[person.stageName] = true;
      if (!person.email) {
        warnings.push('メールアドレス未登録: ' + person.stageName);
      } else if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(person.email)) {
        warnings.push('メールアドレスの形式要確認: ' + person.stageName);
      }
    });

    const judgeSheet = masterSs.getSheetByName('判定表');
    if (judgeSheet) {
      const judgeNames = {};
      judgeSheet.getDataRange().getValues().forEach(function(row) {
        const name = String(row[1] || '').trim();
        if (name && name !== '芸名') judgeNames[name] = true;
      });
      Object.keys(seenNames).forEach(function(name) {
        if (!judgeNames[name]) warnings.push('判定表に未登録: ' + name);
      });
      Object.keys(judgeNames).forEach(function(name) {
        if (!seenNames[name]) warnings.push('外注連絡票に未登録: ' + name);
      });
    }
  } catch (err) {
    issues.push('マスタ確認失敗: ' + err.message);
  }

  try {
    const props = PropertiesService.getScriptProperties();
    const templateId = props.getProperty('TEMPLATE_SPREADSHEET_ID');
    const folderId = props.getProperty('MONTHLY_FOLDER_ID');
    if (!templateId || !folderId) throw new Error('テンプレートまたは月次保存先の設定がありません');
    const templateSs = SpreadsheetApp.openById(templateId);
    const inputSheet = templateSs.getSheetByName('2.入力表');
    if (!inputSheet) throw new Error('テンプレートに「2.入力表」がありません');
    findInputFormulaTemplate_(inputSheet);
    DriveApp.getFolderById(folderId).getName();
  } catch (err) {
    issues.push('月次テンプレート確認失敗: ' + err.message);
  }

  return {
    ok: issues.length === 0,
    issues: issues,
    warnings: warnings,
    jobCount: jobCount,
    dancerCount: dancerCount,
    checkedAt: new Date().toISOString()
  };
}
