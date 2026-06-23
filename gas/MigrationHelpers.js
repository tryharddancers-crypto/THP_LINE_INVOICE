/**
 * Creates THP-owned copies of the Drive resources without changing production settings.
 * Run this first, inspect the copies, then run applyDriveOwnershipMigration.
 */
function stageDriveOwnershipMigration() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('PENDING_MASTER_SPREADSHEET_ID')) {
    throw new Error('移行用コピーは既に作成済みです。作成済みのコピーを確認してください。');
  }

  const masterId = props.getProperty('MASTER_SPREADSHEET_ID');
  const templateId = props.getProperty('TEMPLATE_SPREADSHEET_ID');
  const monthlyFolderId = props.getProperty('MONTHLY_FOLDER_ID');
  if (!masterId || !templateId || !monthlyFolderId) {
    throw new Error('移行元のスクリプトプロパティが不足しています。');
  }

  const stamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  const migrationRoot = DriveApp.createFolder('THP請求書システム_移行データ_' + stamp);
  const masterCopy = DriveApp.getFileById(masterId)
    .makeCopy('請求書マスタ（THP管理用）', migrationRoot);
  const templateCopy = DriveApp.getFileById(templateId)
    .makeCopy('請求書テンプレート（THP管理用）', migrationRoot);
  const monthlyFolderCopy = migrationRoot.createFolder('LINE_INVOICE_月次（THP管理用）');

  copyFolderContentsForMigration_(
    DriveApp.getFolderById(monthlyFolderId),
    monthlyFolderCopy
  );

  props.setProperties({
    PENDING_MASTER_SPREADSHEET_ID: masterCopy.getId(),
    PENDING_TEMPLATE_SPREADSHEET_ID: templateCopy.getId(),
    PENDING_MONTHLY_FOLDER_ID: monthlyFolderCopy.getId(),
    PENDING_MIGRATION_ROOT_FOLDER_ID: migrationRoot.getId()
  }, false);

  Logger.log('移行用コピー作成完了: ' + migrationRoot.getUrl());
  Logger.log('マスタ: ' + masterCopy.getUrl());
  Logger.log('テンプレート: ' + templateCopy.getUrl());
  Logger.log('月次フォルダ: ' + monthlyFolderCopy.getUrl());
}

/** Copies all files and subfolders while retaining the existing hierarchy. */
function copyFolderContentsForMigration_(sourceFolder, destinationFolder) {
  const files = sourceFolder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    file.makeCopy(file.getName(), destinationFolder);
  }

  const folders = sourceFolder.getFolders();
  while (folders.hasNext()) {
    const sourceChild = folders.next();
    const destinationChild = destinationFolder.createFolder(sourceChild.getName());
    copyFolderContentsForMigration_(sourceChild, destinationChild);
  }
}

/** Switches production to the THP-owned copies created by stageDriveOwnershipMigration. */
function applyDriveOwnershipMigration() {
  const props = PropertiesService.getScriptProperties();
  const masterId = props.getProperty('PENDING_MASTER_SPREADSHEET_ID');
  const templateId = props.getProperty('PENDING_TEMPLATE_SPREADSHEET_ID');
  const monthlyFolderId = props.getProperty('PENDING_MONTHLY_FOLDER_ID');
  if (!masterId || !templateId || !monthlyFolderId) {
    throw new Error('先にstageDriveOwnershipMigrationを実行してください。');
  }

  props.setProperties({
    MASTER_SPREADSHEET_ID: masterId,
    OUTSOURCE_MASTER_SPREADSHEET_ID: masterId,
    TEMPLATE_SPREADSHEET_ID: templateId,
    MONTHLY_FOLDER_ID: monthlyFolderId
  }, false);

  const data = getMasterData();
  const currentMonthly = getOrCreateMonthlySpreadsheet(new Date());
  Logger.log(
    'THP所有データへ切替完了: 案件=' + data.jobList.length +
    '件 / 担当者=' + data.dancerNames.length +
    '人 / 当月=' + currentMonthly.getUrl()
  );
}
