/**
 * 物品貸出申請システム (Google Apps Script)
 *
 * 構成:
 * - 申請スプレッドシート: 申請履歴・物品マスタ・採番管理
 * - 日別台帳シート: 日別の貸出数/残数（同一スプレッドシート内）
 *
 * 事前設定:
 * Script Properties に以下を登録
 * - APPLICATION_SHEET_ID: 申請管理スプレッドシートID
 */

const SHEET_NAMES = {
  applications: '申請',
  master: '物品マスタ',
  sequence: '採番管理',
  ledger: '日別台帳',
  managers: '管理者マスタ',
};

const APPLICATION_COLUMNS = {
  timestamp: 1,
  applicationNo: 2,
  seq: 3,
  check: 4,
  department: 5,
  applicant: 6,
  email: 7,
  tel: 8,
  startDate: 9,
  pickupTime: 10,
  endDate: 11,
  returnTime: 12,
  itemCode: 13,
  itemName: 14,
  quantity: 15,
  applicationStatus: 16,
  resolveReason: 17,
  preparationStatus: 18,
  note: 19,
  total: 19,
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('操作')
    .addItem('解決', 'operationResolveOverLimit_')
    .addItem('リセット', 'operationResetSheets_')
    .addToUi();
}

function operationResolveOverLimit_() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('解決理由を入力しでください', '解決理由を入力しでください', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const reason = String(response.getResponseText() || '').trim();
  if (!reason) {
    ui.alert('解決理由を入力しでください。');
    return;
  }

  const sheet = getApplicationSpreadsheet_().getSheetByName(SHEET_NAMES.applications);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const rowCount = lastRow - 1;
  const values = sheet.getRange(2, 1, rowCount, APPLICATION_COLUMNS.total).getValues();
  const resolvedTargets = [];

  let updated = 0;
  values.forEach((row, idx) => {
    const checked = row[APPLICATION_COLUMNS.check - 1] === true;
    const status = String(row[APPLICATION_COLUMNS.applicationStatus - 1] || '').trim();
    if (!checked || status !== '申請数超過') return;

    const rowNumber = idx + 2;
    sheet.getRange(rowNumber, 1, 1, APPLICATION_COLUMNS.quantity).setBackground('#fff2cc');
    sheet.getRange(rowNumber, APPLICATION_COLUMNS.applicationStatus).setValue('解決');
    sheet.getRange(rowNumber, APPLICATION_COLUMNS.resolveReason).setValue(reason);
    resolvedTargets.push({
      startDate: normalizeDate_(row[APPLICATION_COLUMNS.startDate - 1]),
      endDate: normalizeDate_(row[APPLICATION_COLUMNS.endDate - 1]),
      itemCode: String(row[APPLICATION_COLUMNS.itemCode - 1] || '').trim(),
    });
    updated += 1;
  });

  if (updated === 0) {
    ui.alert('チェック済みかつ「申請数超過」の行がありません。');
    return;
  }

  markResolvedRowsOnLedger_(resolvedTargets);
  ui.alert(`${updated}件を解決に更新しました。`);
}

function markResolvedRowsOnLedger_(resolvedTargets) {
  if (!resolvedTargets || resolvedTargets.length === 0) return;

  const ledgerSheet = getLedgerSheet_();
  const masterItems = getMasterItems_();
  const columnMap = getLedgerColumnMap_(ledgerSheet, masterItems);
  const dateRows = getLedgerRowsByDate_(ledgerSheet);

  resolvedTargets.forEach((target) => {
    if (!target.itemCode || !target.startDate || !target.endDate) return;
    const cols = columnMap[target.itemCode];
    if (!cols) return;

    for (const day of iterateDates_(target.startDate, target.endDate)) {
      const row = dateRows[formatDateKey_(day)];
      if (!row) continue;
      ledgerSheet.getRange(row, cols.borrowedCol).setBackground('#fff2cc');
      ledgerSheet.getRange(row, cols.remainingCol).setBackground('#fff2cc');
    }
  });
}

function operationResetSheets_() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert('確認', 'CSVをPCへダウンロードしてから「申請」「日別台帳」の2行目以降を削除します。よろしいですか？', ui.ButtonSet.OK_CANCEL);
  if (confirm !== ui.Button.OK) return;

  const ss = getApplicationSpreadsheet_();
  const appSheet = ss.getSheetByName(SHEET_NAMES.applications);
  const ledgerSheet = ss.getSheetByName(SHEET_NAMES.ledger);
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');

  const payload = {
    applicationFileName: `申請_${timestamp}.csv`,
    ledgerFileName: `日別台帳_${timestamp}.csv`,
    applicationCsv: buildSheetCsv_(appSheet),
    ledgerCsv: buildSheetCsv_(ledgerSheet),
  };

  const html = HtmlService.createHtmlOutput(buildResetDownloadDialogHtml_(payload))
    .setWidth(520)
    .setHeight(300);
  ui.showModalDialog(html, 'CSVダウンロード');
}

function operationResetRowsOnly() {
  const ss = getApplicationSpreadsheet_();
  const appSheet = ss.getSheetByName(SHEET_NAMES.applications);
  const ledgerSheet = ss.getSheetByName(SHEET_NAMES.ledger);

  [appSheet, ledgerSheet].forEach((sheet) => {
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      sheet.deleteRows(2, lastRow - 1);
    }
  });

  return { success: true };
}

function buildResetDownloadDialogHtml_(payload) {
  const data = JSON.stringify(payload).replace(/</g, '\\u003c');
  return `<!DOCTYPE html>
<html lang="ja">
  <head>
    <meta charset="UTF-8" />
    <style>
      body { font-family: Arial, sans-serif; padding: 12px; color: #1f2a44; }
      button { margin-right: 8px; padding: 8px 14px; border: 0; border-radius: 8px; cursor: pointer; color: #fff; background: #4f46e5; }
      button.secondary { background: #5f6f96; }
      .note { font-size: 12px; color: #60719a; margin-bottom: 12px; }
      .ok { margin-top: 12px; color: #0f5132; font-weight: bold; }
    </style>
  </head>
  <body>
    <div class="note">「CSVダウンロードしてリセット」を押すと2ファイルをPCへ保存し、その後シートの2行目以降を削除します。</div>
    <button onclick="runReset()">CSVダウンロードしてリセット</button>
    <button class="secondary" onclick="google.script.host.close()">キャンセル</button>
    <div id="result" class="ok"></div>
    <script>
      const payload = ${data};
      function triggerDownload(fileName, csvText) {
        const blob = new Blob([csvText], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = fileName;
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(url);
      }
      function runReset() {
        triggerDownload(payload.applicationFileName, payload.applicationCsv);
        triggerDownload(payload.ledgerFileName, payload.ledgerCsv);
        google.script.run.withSuccessHandler(() => {
          document.getElementById('result').textContent = 'リセットが完了しました。';
          setTimeout(() => google.script.host.close(), 600);
        }).operationResetRowsOnly();
      }
    </script>
  </body>
</html>`;
}

function buildSheetCsv_(sheet) {
  if (!sheet) return '';
  const values = sheet.getDataRange().getDisplayValues();
  return `\ufeff${values.map((row) => row.map(csvEscape_).join(',')).join('\r\n')}`;
}

function csvEscape_(value) {
  const text = String(value == null ? '' : value);
  if (/[",\n\r]/.test(text)) {
    return `"${text.replace(/"/g, '""')}"`;
  }
  return text;
}



function applyPreparationStatusValidation_(sheet, startRow, rowCount) {
  if (!sheet) return;

  const validation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['準備前', '準備済み', '準備不要'], true)
    .setAllowInvalid(false)
    .build();

  if (startRow && rowCount) {
    const range = sheet.getRange(startRow, APPLICATION_COLUMNS.preparationStatus, rowCount, 1);
    range.setDataValidation(validation);
    const values = range.getValues().map((row) => [row[0] || '準備前']);
    range.setValues(values);
    range.setDataValidation(validation);
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const count = lastRow - 1;
  const range = sheet.getRange(2, APPLICATION_COLUMNS.preparationStatus, count, 1);
  range.setDataValidation(validation);
  const values = range.getValues().map((row) => [row[0] || '準備前']);
  range.setValues(values);
}

function adjustApplicationSheetLayout_(sheet) {
  if (!sheet) return;
  sheet.setColumnWidth(APPLICATION_COLUMNS.check, 90);
  sheet.setColumnWidth(APPLICATION_COLUMNS.applicationStatus, 120);
  sheet.setColumnWidth(APPLICATION_COLUMNS.resolveReason, 180);
  sheet.setColumnWidth(APPLICATION_COLUMNS.preparationStatus, 130);
  sheet.setColumnWidth(APPLICATION_COLUMNS.note, 220);
}

function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || 'form';
  if (page === 'admin') {
    return HtmlService.createTemplateFromFile('Admin')
      .evaluate()
      .setTitle('物品マスタ管理')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('物品貸出申請フォーム')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(fileName) {
  return HtmlService.createHtmlOutputFromFile(fileName).getContent();
}

function initializeSheets() {
  const appSs = getApplicationSpreadsheet_();

  const appHeaders = [
    '登録日時', '申請No', '通番', 'チェック', '申請部署', '申請者氏名', 'メールアドレス', 'Tel',
    '借用開始日', '受取時間', '返却日', '返却時間', '物品コード', '物品名', '数量', '申請状況', '解決理由', '準備状況', '備考'
  ];
  const masterHeaders = ['物品コード', '物品名', '初期在庫', '有効'];
  const sequenceHeaders = ['最新連番'];
  const managerHeaders = ['氏名', 'メールアドレス', '有効'];

  ensureSheetWithHeaders_(appSs, SHEET_NAMES.applications, appHeaders);
  ensureSheetWithHeaders_(appSs, SHEET_NAMES.master, masterHeaders);
  ensureSheetWithHeaders_(appSs, SHEET_NAMES.sequence, sequenceHeaders);
  ensureSheetWithHeaders_(appSs, SHEET_NAMES.managers, managerHeaders);
  const ledgerSheet = getLedgerSheet_();
  ensureLedgerLayout_(ledgerSheet, getMasterItems_());
  applyPreparationStatusValidation_(appSs.getSheetByName(SHEET_NAMES.applications));
  adjustApplicationSheetLayout_(appSs.getSheetByName(SHEET_NAMES.applications));
  applyWorkbookStyle_();
}

function getAvailableItems() {
  const items = getMasterItems_().filter((item) => item.active);
  return items.map((item) => ({
    code: item.code,
    name: item.name,
    maxSelectable: Math.max(Number(item.stock) || 0, 0),
  }));
}

function submitApplication(formData) {
  validateForm_(formData);

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const appSs = getApplicationSpreadsheet_();
    const appSheet = appSs.getSheetByName(SHEET_NAMES.applications);

    const startDate = normalizeDate_(new Date(formData.startDate));
    const endDate = normalizeDate_(new Date(formData.endDate));

    const masterMap = getMasterMap_();

    formData.items.forEach((item) => {
      const masterItem = masterMap[item.code];
      if (!masterItem || !masterItem.active) {
        throw new Error(`無効な物品が指定されました: ${item.code}`);
      }
    });

    const applicationNo = nextApplicationNo_();
    const timestamp = new Date();

    const rows = formData.items.map((item, index) => {
      const masterItem = masterMap[item.code];
      return [
        timestamp,
        `'${applicationNo}`,
        index + 1,
        false,
        formData.department,
        formData.applicant,
        formData.email,
        formData.tel,
        startDate,
        formData.pickupTime,
        endDate,
        formData.returnTime,
        item.code,
        masterItem.name,
        item.quantity,
        '正常',
        '',
        '準備前',
        '',
      ];
    });

    let insertedStartRow = 0;
    if (rows.length > 0) {
      insertedStartRow = appSheet.getLastRow() + 1;
      appSheet.getRange(insertedStartRow, 1, rows.length, rows[0].length).setValues(rows);
      appSheet.getRange(insertedStartRow, APPLICATION_COLUMNS.check, rows.length, 1).insertCheckboxes();
      applyPreparationStatusValidation_(appSheet, insertedStartRow, rows.length);
      updateLedger_(startDate, endDate, formData.items);
    }

    const summary = buildApplicationSummary_(applicationNo, formData, masterMap, startDate, endDate);
    sendApplicationMail_(summary);
    sendAdminNotificationMail_(summary);
    const overLimitAlerts = collectOverLimitAlerts_(applicationNo, formData, masterMap);
    markOverLimitApplicationRows_(appSheet, insertedStartRow, formData.items, overLimitAlerts);
    sendOverLimitAlertMail_(summary, overLimitAlerts);

    return {
      success: true,
      applicationNo,
      message: `申請を受け付けました。申請No: ${applicationNo}`,
      summary,
    };
  } finally {
    lock.releaseLock();
  }
}


function buildApplicationSummary_(applicationNo, formData, masterMap, startDate, endDate) {
  const tz = Session.getScriptTimeZone();
  return {
    applicationNo,
    department: formData.department,
    applicant: formData.applicant,
    email: formData.email,
    tel: formData.tel,
    startDate: Utilities.formatDate(startDate, tz, 'yyyy-MM-dd'),
    pickupTime: formData.pickupTime,
    endDate: Utilities.formatDate(endDate, tz, 'yyyy-MM-dd'),
    returnTime: formData.returnTime,
    items: formData.items.map((item) => ({
      name: masterMap[item.code] ? masterMap[item.code].name : item.code,
      quantity: Number(item.quantity),
    })),
  };
}

function sendApplicationMail_(summary) {
  const itemLines = summary.items
    .map((item) => `・${item.name} × ${item.quantity}`)
    .join('\n');

  const subject = `【物品貸出申請受付】申請No: ${summary.applicationNo}`;
  const body = [
    '物品貸出申請を受け付けました。',
    '',
    `申請No: ${summary.applicationNo}`,
    `申請部署: ${summary.department}`,
    `申請者氏名: ${summary.applicant}`,
    `メールアドレス: ${summary.email}`,
    `Tel: ${summary.tel}`,
    `借用開始日: ${summary.startDate}`,
    `受取時間: ${summary.pickupTime}`,
    `返却日: ${summary.endDate}`,
    `返却時間: ${summary.returnTime}`,
    '',
    '借用物品:',
    itemLines,
  ].join('\n');

MailApp.sendEmail({
    to: summary.email,
    subject,
    body,
    name: '物品貸出申請システム',
  });
}


function sendAdminNotificationMail_(summary) {
  const to = getActiveManagerRecipient_();
  if (!to) return;

  const itemLines = summary.items.map((item) => `・${item.name} × ${item.quantity}`).join('\n');
  const subject = `【管理者通知】新規物品貸出申請 No:${summary.applicationNo}`;
  const body = [
    '新しい物品貸出申請が登録されました。',
    '',
    `申請No: ${summary.applicationNo}`,
    `申請部署: ${summary.department}`,
    `申請者氏名: ${summary.applicant}`,
    `メールアドレス: ${summary.email}`,
    `Tel: ${summary.tel}`,
    `借用開始日: ${summary.startDate}`,
    `受取時間: ${summary.pickupTime}`,
    `返却日: ${summary.endDate}`,
    `返却時間: ${summary.returnTime}`,
    '',
    '借用物品:',
    itemLines,
  ].join('\n');

  MailApp.sendEmail({
    to,
    subject,
    body,
    name: '物品貸出申請システム',
  });
}

function sendOverLimitAlertMail_(summary, alerts) {
  if (!alerts || alerts.length === 0) return;
  const to = getActiveManagerRecipient_();
  if (!to) return;

  const lines = alerts.map((a) => [
    `対象申請: ${summary.applicationNo}`,
    `申請部署: ${summary.department}`,
    `申請者氏名: ${summary.applicant}`,
    `メールアドレス: ${summary.email}`,
    `Tel: ${summary.tel}`,
    `対象日付: ${a.date}`,
    `対象曜日: ${a.weekday}`,
    `対象物品: ${a.itemName}`,
    `超過数量: ${a.excessQty}`,
    `貸出数: ${a.borrowedQty}`,
  ].join('\n')).join('\n\n--------------------\n\n');

  const body = [
    '新規申請により、貸出数が在庫を超過しました。',
    '',
    lines,
  ].join('\n');

  MailApp.sendEmail({
    to,
    subject: `【在庫超過アラート】申請No:${summary.applicationNo}`,
    body,
    name: '物品貸出申請システム',
  });
}

function getActiveManagerRecipient_() {
  const managers = getManagerData().filter((m) => m.active && /@/.test(m.email));
  if (managers.length === 0) return '';
  return managers.map((m) => m.email).join(',');
}

function collectOverLimitAlerts_(applicationNo, formData, masterMap) {
  const appSheet = getApplicationSpreadsheet_().getSheetByName(SHEET_NAMES.applications);
  const appLastRow = appSheet.getLastRow();
  if (appLastRow < 2) return [];

  const appRows = appSheet.getRange(2, 1, appLastRow - 1, APPLICATION_COLUMNS.total).getValues();
  const borrowedBefore = {};
  const borrowedAfter = {};

  appRows.forEach((row) => {
    const rowAppNo = String(row[APPLICATION_COLUMNS.applicationNo - 1] || '').replace(/^'/, '');
    const startDate = normalizeDate_(row[APPLICATION_COLUMNS.startDate - 1]);
    const endDate = normalizeDate_(row[APPLICATION_COLUMNS.endDate - 1]);
    const code = String(row[APPLICATION_COLUMNS.itemCode - 1] || '').trim();
    const qty = Number(row[APPLICATION_COLUMNS.quantity - 1]) || 0;
    if (!code || qty <= 0) return;

    for (const day of iterateDates_(startDate, endDate)) {
      const key = `${formatDateKey_(day)}::${code}`;
      borrowedAfter[key] = (borrowedAfter[key] || 0) + qty;
      if (rowAppNo !== applicationNo) {
        borrowedBefore[key] = (borrowedBefore[key] || 0) + qty;
      }
    }
  });

  const start = normalizeDate_(new Date(formData.startDate));
  const end = normalizeDate_(new Date(formData.endDate));
  const alertMap = {};

  for (const day of iterateDates_(start, end)) {
    const dateKey = formatDateKey_(day);
    formData.items.forEach((item) => {
      const code = String(item.code || '').trim();
      const m = masterMap[code];
      if (!m) return;

      const stock = Number(m.stock) || 0;
      const key = `${dateKey}::${code}`;
      const before = borrowedBefore[key] || 0;
      const after = borrowedAfter[key] || 0;
      if (!(after > stock)) return;

      alertMap[key] = {
        code,
        date: dateKey,
        weekday: getWeekdayJa_(day),
        itemName: m.name,
        excessQty: after - stock,
        borrowedQty: after,
      };
    });
  }

  return Object.keys(alertMap).sort().map((k) => alertMap[k]);
}

function markOverLimitApplicationRows_(appSheet, insertedStartRow, submittedItems, overLimitAlerts) {
  if (!insertedStartRow || !submittedItems || submittedItems.length === 0) return;
  const overLimitCodeSet = (overLimitAlerts || []).reduce((acc, alert) => {
    if (alert && alert.code) acc[alert.code] = true;
    return acc;
  }, {});

  submittedItems.forEach((item, idx) => {
    const rowNumber = insertedStartRow + idx;
    const isOverLimit = Boolean(overLimitCodeSet[item.code]);
    if (isOverLimit) {
      appSheet.getRange(rowNumber, 1, 1, APPLICATION_COLUMNS.quantity).setBackground('#f4cccc');
      appSheet.getRange(rowNumber, APPLICATION_COLUMNS.applicationStatus).setValue('申請数超過');
      return;
    }
    appSheet.getRange(rowNumber, APPLICATION_COLUMNS.applicationStatus).setValue('正常');
  });
}

function getManagerData() {
  const header = ['氏名', 'メールアドレス', '有効'];
  const ss = getApplicationSpreadsheet_();
  ensureSheetWithHeaders_(ss, SHEET_NAMES.managers, header);
  const sheet = ss.getSheetByName(SHEET_NAMES.managers);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const rows = sheet.getRange(2, 1, lastRow - 1, header.length).getValues();
  return rows
    .filter((r) => r[0] && r[1])
    .map((r) => ({
      name: String(r[0]).trim(),
      email: String(r[1]).trim(),
      active: String(r[2]).toUpperCase() !== 'FALSE',
    }));
}

function saveManagerData(managers) {
  const header = ['氏名', 'メールアドレス', '有効'];
  const ss = getApplicationSpreadsheet_();
  ensureSheetWithHeaders_(ss, SHEET_NAMES.managers, header);
  const sheet = ss.getSheetByName(SHEET_NAMES.managers);
  const normalized = (managers || []).map((m) => [
    String(m.name || '').trim(),
    String(m.email || '').trim(),
    m.active ? 'TRUE' : 'FALSE',
  ]).filter((r) => r[0] && r[1]);

  sheet.clearContents();
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  if (normalized.length > 0) {
    sheet.getRange(2, 1, normalized.length, normalized[0].length).setValues(normalized);
  }

  applyWorkbookStyle_();
  return { success: true, message: '管理者を登録しました。' };
}

function getMasterData() {
  return getMasterItems_();
}

function saveMasterData(payload) {
  const items = Array.isArray(payload) ? payload : (payload && payload.items ? payload.items : []);
  const appSs = getApplicationSpreadsheet_();
  const sheet = appSs.getSheetByName(SHEET_NAMES.master);

  const normalized = items.map((item) => [
    item.code,
    item.name,
    Number(item.stock),
    item.active ? 'TRUE' : 'FALSE',
  ]);

  const header = ['物品コード', '物品名', '初期在庫', '有効'];
  sheet.clearContents();
  sheet.getRange(1, 1, 1, header.length).setValues([header]);

  if (normalized.length > 0) {
    sheet.getRange(2, 1, normalized.length, normalized[0].length).setValues(normalized);
  }

  const ledgerSheet = getLedgerSheet_();
  ensureLedgerLayout_(ledgerSheet, getMasterItems_());
  applyPreparationStatusValidation_(appSs.getSheetByName(SHEET_NAMES.applications));
  adjustApplicationSheetLayout_(appSs.getSheetByName(SHEET_NAMES.applications));
  applyWorkbookStyle_();

  return { success: true };
}



function reflectLedgerDateRange(ledgerStartDate, ledgerEndDate) {
  if (!ledgerStartDate || !ledgerEndDate) {
    throw new Error('日別台帳の開始日と終了日を入力してください。');
  }

  const start = normalizeDate_(new Date(ledgerStartDate));
  const end = normalizeDate_(new Date(ledgerEndDate));
  if (start.getTime() > end.getTime()) {
    throw new Error('台帳の日付範囲は開始日 <= 終了日で指定してください。');
  }

  const ledgerSheet = getLedgerSheet_();
  const masterItems = getMasterItems_();
  ensureLedgerLayout_(ledgerSheet, masterItems);
  ensureLedgerDateRows_(ledgerSheet, start, end);
  syncLedgerFromApplications_(ledgerSheet, masterItems);
  applyWorkbookStyle_();

  return {
    success: true,
    message: `日別台帳に ${Utilities.formatDate(start, Session.getScriptTimeZone(), 'yyyy-MM-dd')} ～ ${Utilities.formatDate(end, Session.getScriptTimeZone(), 'yyyy-MM-dd')} を反映しました。`,
  };
}

function applyWorkbookStyle_() {
  const ss = getApplicationSpreadsheet_();
  [
    ss.getSheetByName(SHEET_NAMES.applications),
    ss.getSheetByName(SHEET_NAMES.master),
    ss.getSheetByName(SHEET_NAMES.sequence),
    ss.getSheetByName(SHEET_NAMES.ledger),
    ss.getSheetByName(SHEET_NAMES.managers),
  ].forEach((sheet) => {
    if (!sheet) return;
    applySheetHeaderStyle_(sheet);
    applySheetBodyStyle_(sheet);
  });

  const ledger = ss.getSheetByName(SHEET_NAMES.ledger);
  if (ledger) {
    ledger.setFrozenRows(1);
    ledger.getRange('A:A').setNumberFormat('yyyy-mm-dd');
    if (ledger.getLastColumn() >= 2) {
      ledger.getRange('B:B').setHorizontalAlignment('center');
    }
  }

  const applications = ss.getSheetByName(SHEET_NAMES.applications);
  if (applications) {
    applications.setFrozenRows(1);
    applications.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm');
    applications.getRange('B:B').setNumberFormat('@');
    applications.getRange('I:I').setNumberFormat('yyyy-mm-dd');
    applications.getRange('K:K').setNumberFormat('yyyy-mm-dd');
    applyPreparationStatusValidation_(applications);
    adjustApplicationSheetLayout_(applications);
  }
}

function applySheetHeaderStyle_(sheet) {
  const lastCol = Math.max(1, sheet.getLastColumn());
  const header = sheet.getRange(1, 1, 1, lastCol);
  header
    .setBackground('#34495e')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontFamily('Noto Sans JP');
}

function applySheetBodyStyle_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(1, sheet.getLastColumn());
  if (lastRow >= 2) {
    const body = sheet.getRange(2, 1, lastRow - 1, lastCol);
    body
      .setFontFamily('Noto Sans JP')
      .setFontColor('#1f2a44')
      .setBackground('#ffffff')
      .setVerticalAlignment('middle');
  }

  sheet.autoResizeColumns(1, lastCol);
}

function ensureSheetWithHeaders_(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  const currentHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const mismatch = headers.some((h, i) => currentHeaders[i] !== h);
  if (mismatch) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}

function getApplicationSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('APPLICATION_SHEET_ID');
  if (!id) {
    throw new Error('Script Properties に APPLICATION_SHEET_ID を設定してください。');
  }
  return SpreadsheetApp.openById(id);
}

function getMasterItems_() {
  const appSs = getApplicationSpreadsheet_();
  const sheet = appSs.getSheetByName(SHEET_NAMES.master);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const rows = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  return rows
    .filter((r) => r[0] && r[1])
    .map((r) => ({
      code: String(r[0]).trim(),
      name: String(r[1]).trim(),
      stock: Number(r[2]) || 0,
      active: String(r[3]).toUpperCase() !== 'FALSE',
    }));
}

function getMasterMap_() {
  return getMasterItems_().reduce((acc, item) => {
    acc[item.code] = item;
    return acc;
  }, {});
}

function getMinRemainingInRange_(item, startDate, endDate) {
  const ledgerSheet = getLedgerSheet_();
  const columnMap = ensureLedgerColumns_(ledgerSheet, getMasterItems_());
  const itemCols = columnMap[item.code];
  if (!itemCols) return item.stock;

  const dateRows = getLedgerRowsByDate_(ledgerSheet);
  let minRemaining = item.stock;

  for (const day of iterateDates_(startDate, endDate)) {
    const dateKey = formatDateKey_(day);
    const row = dateRows[dateKey];
    if (!row) {
      minRemaining = Math.min(minRemaining, item.stock);
      continue;
    }
    const remaining = getCellNumber_(ledgerSheet, row, itemCols.remainingCol, item.stock);
    minRemaining = Math.min(minRemaining, remaining);
  }

  return minRemaining;
}


function nextApplicationNo_() {
  const appSs = getApplicationSpreadsheet_();
  const seqSheet = appSs.getSheetByName(SHEET_NAMES.sequence);

  const lastRow = seqSheet.getLastRow();
  let current = 0;

  if (lastRow >= 2) {
    const values = seqSheet.getRange(2, 1, lastRow - 1, 2).getValues();
    values.forEach((row) => {
      const c1 = Number(row[0]) || 0;
      const c2 = Number(row[1]) || 0;
      current = Math.max(current, c1, c2);
    });
  }

  const next = current + 1;
  seqSheet.getRange(1, 1).setValue('最新連番');
  seqSheet.getRange(2, 1).setValue(next);
  if (seqSheet.getLastColumn() >= 2) {
    seqSheet.getRange(2, 2, Math.max(1, seqSheet.getLastRow() - 1), 1).clearContent();
  }

  return String(next).padStart(5, '0');
}

function getFiscalYear_(date) {
  const y = date.getFullYear();
  const m = date.getMonth() + 1;
  return m >= 4 ? y : y - 1;
}

function updateLedger_(startDate, endDate, items) {
  const ledgerSheet = getLedgerSheet_();
  const masterItems = getMasterItems_();

  ensureLedgerLayout_(ledgerSheet, masterItems);
  ensureLedgerDateRows_(ledgerSheet, startDate, endDate);
  syncLedgerFromApplications_(ledgerSheet, masterItems);
}

function getLedgerSheet_() {
  const ss = getApplicationSpreadsheet_();
  let sheet = ss.getSheetByName(SHEET_NAMES.ledger);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.ledger);
  }
  return sheet;
}

function ensureLedgerLayout_(sheet, masterItems) {
  ensureLedgerColumns_(sheet, masterItems);
}

function ensureLedgerColumns_(sheet, masterItems) {
  let currentLastColumn = sheet.getLastColumn();
  if (currentLastColumn < 1) {
    sheet.insertColumnBefore(1);
    currentLastColumn = 1;
  }

  const currentHeader = sheet.getRange(1, 1, 1, currentLastColumn).getValues()[0];

  const expectedHeader = buildLedgerHeader_(masterItems);
  const needsResize = currentLastColumn < expectedHeader.length;
  if (needsResize) {
    sheet.insertColumnsAfter(currentLastColumn, expectedHeader.length - currentLastColumn);
    currentLastColumn = expectedHeader.length;
  }

  expectedHeader.forEach((header, idx) => {
    const col = idx + 1;
    if (currentHeader[idx] !== header) {
      sheet.getRange(1, col).setValue(header);
    }
  });

  sheet.autoResizeColumns(1, expectedHeader.length);
  if (expectedHeader.length >= 3) {
    for (let col = 3; col <= expectedHeader.length; col += 1) {
      sheet.setColumnWidth(col, Math.max(sheet.getColumnWidth(col), 120));
    }
  }

  return getLedgerColumnMap_(sheet, masterItems);
}

function buildLedgerHeader_(masterItems) {
  const headers = ['日付', '曜日'];
  masterItems.forEach((item) => {
    headers.push(`${item.name}_貸出数`);
    headers.push(`${item.name}_残数`);
  });
  return headers;
}

function getLedgerColumnMap_(sheet, masterItems) {
  const map = {};
  masterItems.forEach((item, idx) => {
    const borrowedCol = 3 + idx * 2;
    const remainingCol = borrowedCol + 1;
    map[item.code] = { borrowedCol, remainingCol };
  });
  return map;
}

function getLedgerRowsByDate_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const map = {};
  values.forEach((v, idx) => {
    if (!v[0]) return;
    const key = formatDateKey_(normalizeDate_(v[0]));
    map[key] = idx + 2;
  });
  return map;
}

function ensureLedgerDateRows_(sheet, startDate, endDate) {
  const dateRows = getLedgerRowsByDate_(sheet);
  const inserts = [];

  for (const day of iterateDates_(startDate, endDate)) {
    const key = formatDateKey_(day);
    if (!dateRows[key]) {
      inserts.push([new Date(day.getTime()), getWeekdayJa_(day)]);
    }
  }

  if (inserts.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, inserts.length, 2).setValues(inserts);
  }

  return getLedgerRowsByDate_(sheet);
}

function getCellNumber_(sheet, row, col, fallback) {
  const value = sheet.getRange(row, col).getValue();
  return value === '' || value == null ? Number(fallback) : Number(value) || 0;
}

function formatDateKey_(date) {
  return Utilities.formatDate(normalizeDate_(date), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}



function syncLedgerFromApplications_(ledgerSheet, masterItems) {
  ensureLedgerLayout_(ledgerSheet, masterItems);

  const dateRows = getLedgerRowsByDate_(ledgerSheet);
  const columnMap = getLedgerColumnMap_(ledgerSheet, masterItems);
  const masterMap = getMasterMap_();

  const lastRow = ledgerSheet.getLastRow();
  const lastCol = ledgerSheet.getLastColumn();
  if (lastRow >= 2 && lastCol >= 3) {
    ledgerSheet.getRange(2, 3, lastRow - 1, lastCol - 2).clearContent().setBackground(null);
  }

  const appSheet = getApplicationSpreadsheet_().getSheetByName(SHEET_NAMES.applications);
  const appLastRow = appSheet.getLastRow();
  if (appLastRow < 2) return;

  const appRows = appSheet.getRange(2, 1, appLastRow - 1, APPLICATION_COLUMNS.total).getValues();
  const borrowedMap = {};

  appRows.forEach((row) => {
    const startDate = normalizeDate_(row[APPLICATION_COLUMNS.startDate - 1]);
    const endDate = normalizeDate_(row[APPLICATION_COLUMNS.endDate - 1]);
    const code = String(row[APPLICATION_COLUMNS.itemCode - 1]).trim();
    const qty = Number(row[APPLICATION_COLUMNS.quantity - 1]) || 0;
    if (!code || qty <= 0) return;

    for (const day of iterateDates_(startDate, endDate)) {
      const key = `${formatDateKey_(day)}::${code}`;
      borrowedMap[key] = (borrowedMap[key] || 0) + qty;
    }
  });

  Object.keys(dateRows).forEach((dateKey) => {
    const rowNumber = dateRows[dateKey];
    masterItems.forEach((item) => {
      const cols = columnMap[item.code];
      if (!cols) return;
      const borrowed = borrowedMap[`${dateKey}::${item.code}`] || 0;
      const stock = Number(masterMap[item.code] ? masterMap[item.code].stock : item.stock) || 0;
      const remaining = stock - borrowed;
      ledgerSheet.getRange(rowNumber, cols.borrowedCol).setValue(borrowed);
      ledgerSheet.getRange(rowNumber, cols.remainingCol).setValue(remaining);
      setLedgerAlertColor_(ledgerSheet, rowNumber, cols, remaining < 0);
    });
  });
}

function setLedgerAlertColor_(sheet, row, cols, isAlert) {
  const color = isAlert ? '#f4cccc' : null;
  sheet.getRange(row, cols.borrowedCol).setBackground(color);
  sheet.getRange(row, cols.remainingCol).setBackground(color);
}

function getWeekdayJa_(date) {
  return ['日', '月', '火', '水', '木', '金', '土'][normalizeDate_(date).getDay()];
}

function* iterateDates_(startDate, endDate) {
  const day = new Date(startDate.getTime());
  while (day.getTime() <= endDate.getTime()) {
    yield new Date(day.getTime());
    day.setDate(day.getDate() + 1);
  }
}

function normalizeDate_(date) {
  const d = new Date(date);
  d.setHours(0, 0, 0, 0);
  return d;
}

function validateForm_(formData) {
  const required = [
    'department', 'applicant', 'email', 'tel',
    'startDate', 'pickupTime', 'endDate', 'returnTime', 'items'
  ];

  required.forEach((key) => {
    if (!formData[key] || (Array.isArray(formData[key]) && formData[key].length === 0)) {
      throw new Error(`必須項目が未入力です: ${key}`);
    }
  });

  if (!/@/.test(formData.email)) {
    throw new Error('メールアドレスの形式が不正です。');
  }

  const start = normalizeDate_(new Date(formData.startDate));
  const end = normalizeDate_(new Date(formData.endDate));
  if (start.getTime() > end.getTime()) {
    throw new Error('借用開始日は返却日以前を指定してください。');
  }

  formData.items.forEach((item) => {
    if (!item.code || Number(item.quantity) <= 0) {
      throw new Error('借用物品の選択内容が不正です。');
    }
  });
}
