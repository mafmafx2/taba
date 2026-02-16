/**
 * 物品貸出申請システム - Google Apps Script
 *
 * 【セットアップ手順】
 * 1. Googleスプレッドシートを新規作成
 * 2. 「拡張機能」→「Apps Script」を開く
 * 3. このファイルの内容をコードエディタに貼り付け
 * 4. FormPage.html, AdminPage.html も同様にHTMLファイルとして追加
 * 5. setupSheets() を一度実行してシートを初期化
 * 6. 「デプロイ」→「新しいデプロイ」→ ウェブアプリとしてデプロイ
 */

// ===== シート名定数 =====
var SHEET_MASTER = '物品マスタ';
var SHEET_APPLICATIONS = '申請データ';
var SHEET_LEDGER = '貸出台帳';
var SHEET_SETTINGS = '設定';

// ===== ウェブアプリ エントリーポイント =====

/**
 * GETリクエストのハンドラ
 * ?page=admin で管理画面、それ以外は申請フォームを表示
 */
function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'form';

  if (page === 'admin') {
    return HtmlService.createTemplateFromFile('AdminPage')
      .evaluate()
      .setTitle('物品貸出管理 - 管理画面')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return HtmlService.createTemplateFromFile('FormPage')
    .evaluate()
    .setTitle('物品借用申請フォーム')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * HTMLテンプレートでファイルをインクルードするためのヘルパー
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ===== スプレッドシート取得 =====

/**
 * スプレッドシートを取得（コンテナバインドスクリプト前提）
 */
function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

// ===== 初期セットアップ =====

/**
 * シートの初期セットアップ（最初に一度だけ実行）
 */
function setupSheets() {
  var ss = getSpreadsheet();

  // 物品マスタシート
  var masterSheet = ss.getSheetByName(SHEET_MASTER);
  if (!masterSheet) {
    masterSheet = ss.insertSheet(SHEET_MASTER);
    masterSheet.getRange(1, 1, 1, 5).setValues([
      ['物品ID', '物品名', '写真URL', '在庫数', '有効']
    ]);
    masterSheet.getRange(1, 1, 1, 5).setFontWeight('bold');
    masterSheet.setFrozenRows(1);
  }

  // 申請データシート
  var appSheet = ss.getSheetByName(SHEET_APPLICATIONS);
  if (!appSheet) {
    appSheet = ss.insertSheet(SHEET_APPLICATIONS);
    appSheet.getRange(1, 1, 1, 15).setValues([
      ['申請No.', '通番', '申請日時', '申請部署', '申請者氏名',
       'メールアドレス', 'Tel', '借用開始日', '受取時間',
       '返却日', '返却時間', '物品ID', '物品名', '数量', 'ステータス']
    ]);
    appSheet.getRange(1, 1, 1, 15).setFontWeight('bold');
    appSheet.setFrozenRows(1);
  }

  // 設定シート
  var settingsSheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet(SHEET_SETTINGS);
    settingsSheet.getRange(1, 1, 3, 2).setValues([
      ['キー', '値'],
      ['最終申請番号_年度', ''],
      ['最終申請番号_連番', '0']
    ]);
    settingsSheet.getRange(1, 1, 1, 2).setFontWeight('bold');
  }

  // 貸出台帳シート
  var ledgerSheet = ss.getSheetByName(SHEET_LEDGER);
  if (!ledgerSheet) {
    ledgerSheet = ss.insertSheet(SHEET_LEDGER);
  }

  // デフォルトのSheet1を削除（存在する場合）
  var defaultSheet = ss.getSheetByName('Sheet1') || ss.getSheetByName('シート1');
  if (defaultSheet && ss.getSheets().length > 1) {
    try { ss.deleteSheet(defaultSheet); } catch (e) { /* 無視 */ }
  }

  SpreadsheetApp.flush();
  return '初期セットアップが完了しました。';
}

// ===== 年度計算 =====

/**
 * 日付から年度を取得（4月始まり）
 * @param {Date} date
 * @return {number} 年度（例: 2026）
 */
function getFiscalYear(date) {
  var d = new Date(date);
  var year = d.getFullYear();
  var month = d.getMonth() + 1; // 0-indexed → 1-indexed
  if (month < 4) {
    return year - 1;
  }
  return year;
}

// ===== 申請番号生成 =====

/**
 * 次の申請番号を生成（排他制御付き）
 * 形式: YYYYNNNNN（9桁）
 * @return {string} 申請番号
 */
function generateApplicationNo() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    throw new Error('申請番号の生成に失敗しました。しばらくしてから再度お試しください。');
  }

  try {
    var ss = getSpreadsheet();
    var settingsSheet = ss.getSheetByName(SHEET_SETTINGS);
    var data = settingsSheet.getRange(2, 1, 2, 2).getValues();

    var currentFiscalYear = getFiscalYear(new Date());
    var storedFiscalYear = data[0][1];
    var lastSeqNo = parseInt(data[1][1], 10) || 0;

    var newSeqNo;
    if (storedFiscalYear === currentFiscalYear || storedFiscalYear === String(currentFiscalYear)) {
      newSeqNo = lastSeqNo + 1;
    } else {
      // 新年度: リセット
      newSeqNo = 1;
      settingsSheet.getRange(2, 2).setValue(currentFiscalYear);
    }

    settingsSheet.getRange(3, 2).setValue(newSeqNo);
    SpreadsheetApp.flush();

    // 9桁の申請番号を生成: 年度4桁 + 連番5桁
    var seqStr = ('00000' + newSeqNo).slice(-5);
    return String(currentFiscalYear) + seqStr;
  } finally {
    lock.releaseLock();
  }
}

// ===== 物品マスタ管理 =====

/**
 * 有効な物品一覧を取得
 * @return {Array} 物品リスト [{id, name, photoUrl, stock, active}]
 */
function getActiveItems() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_MASTER);
  var lastRow = sheet.getLastRow();

  if (lastRow <= 1) return [];

  var data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  var items = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (row[4] === true || row[4] === 'TRUE' || row[4] === '有効') {
      items.push({
        id: row[0],
        name: row[1],
        photoUrl: convertDriveUrl(row[2]),
        stock: parseInt(row[3], 10) || 0,
        active: true
      });
    }
  }

  return items;
}

/**
 * 全物品一覧を取得（管理画面用）
 * @return {Array} 物品リスト
 */
function getAllItems() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_MASTER);
  var lastRow = sheet.getLastRow();

  if (lastRow <= 1) return [];

  var data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  var items = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    items.push({
      id: row[0],
      name: row[1],
      photoUrl: row[2],
      stock: parseInt(row[3], 10) || 0,
      active: row[4] === true || row[4] === 'TRUE' || row[4] === '有効',
      rowIndex: i + 2
    });
  }

  return items;
}

/**
 * 物品を追加
 * @param {string} name - 物品名
 * @param {string} photoUrl - 写真URL
 * @param {number} stock - 在庫数（貸出上限）
 * @return {Object} 追加結果
 */
function addItem(name, photoUrl, stock) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_MASTER);

  // 次の物品IDを生成
  var lastRow = sheet.getLastRow();
  var newId;
  if (lastRow <= 1) {
    newId = 'ITEM001';
  } else {
    var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    var maxNum = 0;
    for (var i = 0; i < ids.length; i++) {
      var idStr = String(ids[i][0]);
      var match = idStr.match(/ITEM(\d+)/);
      if (match) {
        var num = parseInt(match[1], 10);
        if (num > maxNum) maxNum = num;
      }
    }
    newId = 'ITEM' + ('000' + (maxNum + 1)).slice(-3);
  }

  sheet.appendRow([newId, name, photoUrl || '', parseInt(stock, 10) || 0, '有効']);

  return { success: true, id: newId, message: '物品「' + name + '」を追加しました。' };
}

/**
 * 物品を更新
 * @param {number} rowIndex - 行番号
 * @param {string} name - 物品名
 * @param {string} photoUrl - 写真URL
 * @param {number} stock - 在庫数
 * @param {boolean} active - 有効フラグ
 * @return {Object} 更新結果
 */
function updateItem(rowIndex, name, photoUrl, stock, active) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_MASTER);

  sheet.getRange(rowIndex, 2).setValue(name);
  sheet.getRange(rowIndex, 3).setValue(photoUrl || '');
  sheet.getRange(rowIndex, 4).setValue(parseInt(stock, 10) || 0);
  sheet.getRange(rowIndex, 5).setValue(active ? '有効' : '無効');

  return { success: true, message: '物品「' + name + '」を更新しました。' };
}

/**
 * 物品を無効化（論理削除）
 * @param {number} rowIndex - 行番号
 * @return {Object} 削除結果
 */
function deactivateItem(rowIndex) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_MASTER);
  sheet.getRange(rowIndex, 5).setValue('無効');
  return { success: true, message: '物品を無効化しました。' };
}

/**
 * Google DriveのURLを表示可能なURLに変換
 * @param {string} url
 * @return {string}
 */
function convertDriveUrl(url) {
  if (!url) return '';

  // Google Drive共有URL → 直接表示URLに変換
  var match = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (match) {
    return 'https://lh3.googleusercontent.com/d/' + match[1];
  }

  // 既にID形式で入力された場合
  if (/^[a-zA-Z0-9_-]{20,}$/.test(url)) {
    return 'https://lh3.googleusercontent.com/d/' + url;
  }

  // そのまま返す（外部URL等）
  return url;
}

// ===== 在庫・貸出数計算 =====

/**
 * 指定日に特定物品がいくつ貸出中かを計算
 * @param {string} itemId - 物品ID
 * @param {Date} targetDate - 対象日
 * @return {number} 貸出数
 */
function getLentQuantityOnDate(itemId, targetDate) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_APPLICATIONS);
  var lastRow = sheet.getLastRow();

  if (lastRow <= 1) return 0;

  var data = sheet.getRange(2, 1, lastRow - 1, 15).getValues();
  var target = new Date(targetDate);
  target.setHours(0, 0, 0, 0);

  var totalLent = 0;

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var rowItemId = row[11]; // 物品ID（L列）
    var status = row[14];    // ステータス（O列）

    // キャンセル済みは除外
    if (status === 'キャンセル') continue;

    if (String(rowItemId) !== String(itemId)) continue;

    var startDate = new Date(row[7]); // 借用開始日（H列）
    var endDate = new Date(row[9]);   // 返却日（J列）
    startDate.setHours(0, 0, 0, 0);
    endDate.setHours(0, 0, 0, 0);

    if (target >= startDate && target <= endDate) {
      totalLent += parseInt(row[13], 10) || 0; // 数量（N列）
    }
  }

  return totalLent;
}

/**
 * 指定期間における物品の最小利用可能数を計算
 * @param {string} itemId - 物品ID
 * @param {number} totalStock - 総在庫数
 * @param {string} startDateStr - 開始日（YYYY-MM-DD）
 * @param {string} endDateStr - 終了日（YYYY-MM-DD）
 * @return {number} 利用可能数（期間中の最小値）
 */
function getMinAvailableQuantity(itemId, totalStock, startDateStr, endDateStr) {
  var startDate = new Date(startDateStr);
  var endDate = new Date(endDateStr);
  startDate.setHours(0, 0, 0, 0);
  endDate.setHours(0, 0, 0, 0);

  var minAvailable = totalStock;
  var current = new Date(startDate);

  while (current <= endDate) {
    var lent = getLentQuantityOnDate(itemId, current);
    var available = totalStock - lent;
    if (available < minAvailable) {
      minAvailable = available;
    }
    current.setDate(current.getDate() + 1);
  }

  return Math.max(0, minAvailable);
}

/**
 * フォーム用: 選択可能な物品リスト（利用可能数付き）を取得
 * @param {string} startDateStr - 借用開始日（YYYY-MM-DD）
 * @param {string} endDateStr - 返却日（YYYY-MM-DD）
 * @return {Array} 物品リスト [{id, name, photoUrl, stock, available}]
 */
function getAvailableItems(startDateStr, endDateStr) {
  var items = getActiveItems();

  if (!startDateStr || !endDateStr) {
    // 日付未選択時は在庫数をそのまま返す
    for (var i = 0; i < items.length; i++) {
      items[i].available = items[i].stock;
    }
    return items;
  }

  for (var i = 0; i < items.length; i++) {
    items[i].available = getMinAvailableQuantity(
      items[i].id, items[i].stock, startDateStr, endDateStr
    );
  }

  return items;
}

// ===== 申請処理 =====

/**
 * 申請を送信
 * @param {Object} formData - フォームデータ
 * @return {Object} 送信結果
 */
function submitApplication(formData) {
  // バリデーション
  if (!formData.department || !formData.applicantName || !formData.email ||
      !formData.tel || !formData.startDate || !formData.pickupTime ||
      !formData.returnDate || !formData.returnTime) {
    return { success: false, message: '必須項目をすべて入力してください。' };
  }

  if (!formData.items || formData.items.length === 0) {
    return { success: false, message: '借用物品を1つ以上選択してください。' };
  }

  // 日付チェック
  var startDate = new Date(formData.startDate);
  var returnDate = new Date(formData.returnDate);
  if (returnDate < startDate) {
    return { success: false, message: '返却日は借用開始日以降の日付を選択してください。' };
  }

  var today = new Date();
  today.setHours(0, 0, 0, 0);
  if (startDate < today) {
    return { success: false, message: '借用開始日は本日以降の日付を選択してください。' };
  }

  // 在庫チェック（サーバーサイドで再検証）
  var activeItems = getActiveItems();
  var itemMap = {};
  for (var i = 0; i < activeItems.length; i++) {
    itemMap[activeItems[i].id] = activeItems[i];
  }

  for (var i = 0; i < formData.items.length; i++) {
    var reqItem = formData.items[i];
    var masterItem = itemMap[reqItem.id];

    if (!masterItem) {
      return { success: false, message: '物品「' + reqItem.name + '」は現在利用できません。' };
    }

    var available = getMinAvailableQuantity(
      reqItem.id, masterItem.stock, formData.startDate, formData.returnDate
    );

    if (parseInt(reqItem.quantity, 10) > available) {
      return {
        success: false,
        message: '物品「' + reqItem.name + '」の残数が不足しています（残数: ' + available + '）。'
      };
    }
  }

  // 申請番号を生成
  var applicationNo = generateApplicationNo();
  var now = new Date();
  var timestamp = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');

  // スプレッドシートに記録
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_APPLICATIONS);

  var rows = [];
  for (var i = 0; i < formData.items.length; i++) {
    var item = formData.items[i];
    rows.push([
      applicationNo,
      i + 1,                          // 通番
      timestamp,                      // 申請日時
      formData.department,            // 申請部署
      formData.applicantName,         // 申請者氏名
      formData.email,                 // メールアドレス
      formData.tel,                   // Tel
      formData.startDate,             // 借用開始日
      formData.pickupTime,            // 受取時間
      formData.returnDate,            // 返却日
      formData.returnTime,            // 返却時間
      item.id,                        // 物品ID
      item.name,                      // 物品名
      parseInt(item.quantity, 10),    // 数量
      '申請済'                         // ステータス
    ]);
  }

  // 一括書き込み
  var startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, 15).setValues(rows);

  // 貸出台帳を更新
  updateLedger();

  return {
    success: true,
    applicationNo: applicationNo,
    message: '申請が完了しました。申請No.: ' + applicationNo
  };
}

// ===== 貸出台帳 =====

/**
 * 貸出台帳シートを更新
 * 当月から2ヶ月先までの日別・物品別の貸出数・残数を表示
 */
function updateLedger() {
  var ss = getSpreadsheet();
  var ledgerSheet = ss.getSheetByName(SHEET_LEDGER);
  var items = getActiveItems();

  if (items.length === 0) {
    ledgerSheet.clear();
    ledgerSheet.getRange(1, 1).setValue('有効な物品が登録されていません。');
    return;
  }

  // 表示期間: 今日から60日間
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var endPeriod = new Date(today);
  endPeriod.setDate(endPeriod.getDate() + 60);

  // ヘッダー行を作成
  var headers = ['日付'];
  for (var i = 0; i < items.length; i++) {
    headers.push(items[i].name + ' (貸出)');
    headers.push(items[i].name + ' (残数)');
  }

  // 全申請データを一括取得（パフォーマンス改善）
  var appSheet = ss.getSheetByName(SHEET_APPLICATIONS);
  var appLastRow = appSheet.getLastRow();
  var allAppData = [];
  if (appLastRow > 1) {
    allAppData = appSheet.getRange(2, 1, appLastRow - 1, 15).getValues();
  }

  // 日別データを作成
  var dataRows = [];
  var current = new Date(today);

  while (current <= endPeriod) {
    var row = [Utilities.formatDate(current, 'Asia/Tokyo', 'yyyy/MM/dd')];

    for (var i = 0; i < items.length; i++) {
      var lent = calcLentFromData(allAppData, items[i].id, current);
      var remaining = items[i].stock - lent;
      row.push(lent);
      row.push(Math.max(0, remaining));
    }

    dataRows.push(row);
    current.setDate(current.getDate() + 1);
  }

  // シートをクリアして書き込み
  ledgerSheet.clear();

  // ヘッダー
  ledgerSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  ledgerSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  ledgerSheet.setFrozenRows(1);

  // データ
  if (dataRows.length > 0) {
    ledgerSheet.getRange(2, 1, dataRows.length, headers.length).setValues(dataRows);
  }

  // 残数0のセルを赤背景にする
  for (var r = 0; r < dataRows.length; r++) {
    for (var c = 0; c < items.length; c++) {
      var colIndex = 2 + c * 2 + 1; // 残数列（1-indexed）
      var cellValue = dataRows[r][colIndex - 1];
      if (cellValue <= 0) {
        ledgerSheet.getRange(r + 2, colIndex).setBackground('#ffcccc');
      }
    }
  }
}

/**
 * 全申請データから指定日の貸出数を計算（台帳更新用）
 * @param {Array} allData - 全申請データ
 * @param {string} itemId - 物品ID
 * @param {Date} targetDate - 対象日
 * @return {number} 貸出数
 */
function calcLentFromData(allData, itemId, targetDate) {
  var target = new Date(targetDate);
  target.setHours(0, 0, 0, 0);
  var totalLent = 0;

  for (var i = 0; i < allData.length; i++) {
    var row = allData[i];
    if (row[14] === 'キャンセル') continue;
    if (String(row[11]) !== String(itemId)) continue;

    var startDate = new Date(row[7]);
    var endDate = new Date(row[9]);
    startDate.setHours(0, 0, 0, 0);
    endDate.setHours(0, 0, 0, 0);

    if (target >= startDate && target <= endDate) {
      totalLent += parseInt(row[13], 10) || 0;
    }
  }

  return totalLent;
}

// ===== 管理画面用: 申請データ取得 =====

/**
 * 申請データ一覧を取得
 * @param {number} limit - 取得件数（デフォルト100）
 * @return {Array} 申請データリスト
 */
function getApplications(limit) {
  var maxRows = limit || 100;
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_APPLICATIONS);
  var lastRow = sheet.getLastRow();

  if (lastRow <= 1) return [];

  var startRow = Math.max(2, lastRow - maxRows + 1);
  var numRows = lastRow - startRow + 1;
  var data = sheet.getRange(startRow, 1, numRows, 15).getValues();

  var applications = [];
  for (var i = data.length - 1; i >= 0; i--) {
    var row = data[i];
    applications.push({
      applicationNo: row[0],
      subNo: row[1],
      timestamp: row[2],
      department: row[3],
      applicantName: row[4],
      email: row[5],
      tel: row[6],
      startDate: row[7],
      pickupTime: row[8],
      returnDate: row[9],
      returnTime: row[10],
      itemId: row[11],
      itemName: row[12],
      quantity: row[13],
      status: row[14]
    });
  }

  return applications;
}

/**
 * 申請ステータスを更新
 * @param {string} applicationNo - 申請番号
 * @param {string} newStatus - 新しいステータス
 * @return {Object} 更新結果
 */
function updateApplicationStatus(applicationNo, newStatus) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_APPLICATIONS);
  var lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return { success: false, message: '申請データが見つかりません。' };
  }

  var data = sheet.getRange(2, 1, lastRow - 1, 15).getValues();
  var updated = false;

  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(applicationNo)) {
      sheet.getRange(i + 2, 15).setValue(newStatus);
      updated = true;
    }
  }

  if (updated) {
    updateLedger();
    return { success: true, message: '申請No.' + applicationNo + 'のステータスを「' + newStatus + '」に更新しました。' };
  }

  return { success: false, message: '申請No.' + applicationNo + 'が見つかりません。' };
}

// ===== 写真アップロード =====

/**
 * Base64エンコードされた画像をGoogle Driveに保存
 * @param {string} base64Data - Base64データ（data:image/...;base64,...）
 * @param {string} fileName - ファイル名
 * @return {Object} アップロード結果
 */
function uploadImageToDrive(base64Data, fileName) {
  try {
    // data:image/xxx;base64, の部分を除去
    var parts = base64Data.split(',');
    var mimeMatch = parts[0].match(/data:(.*?);/);
    var mimeType = mimeMatch ? mimeMatch[1] : 'image/png';
    var decoded = Utilities.base64Decode(parts[1]);
    var blob = Utilities.newBlob(decoded, mimeType, fileName);

    // 保存フォルダを取得または作成
    var folderName = '物品貸出_写真';
    var folders = DriveApp.getFoldersByName(folderName);
    var folder;
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(folderName);
    }

    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return {
      success: true,
      fileId: file.getId(),
      url: 'https://drive.google.com/file/d/' + file.getId() + '/view',
      displayUrl: 'https://lh3.googleusercontent.com/d/' + file.getId()
    };
  } catch (e) {
    return { success: false, message: '画像のアップロードに失敗しました: ' + e.message };
  }
}
