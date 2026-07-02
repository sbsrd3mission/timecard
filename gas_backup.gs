/**
 * タイムカード サーバーサイド(Google Apps Script)
 * スプレッドシートをDBとして使用します。
 */

// ===== 全件取得キャッシュ =====
// 全シート読み込み（5〜6秒）の結果を短時間だけサーバー側に記憶して応答を高速化する。
// 打刻・削除などの書き込みがあった瞬間に破棄するため、古いデータが返り続けることはない。
const RECORDS_CACHE_KEY = 'getAll_cache_v1';
const RECORDS_CACHE_SECONDS = 15;

function doGet(e) {
  const action = e.parameter.action;
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (action === 'getAll') {
    try {
      const cached = CacheService.getScriptCache().get(RECORDS_CACHE_KEY);
      if (cached) {
        return ContentService.createTextOutput(cached)
          .setMimeType(ContentService.MimeType.JSON);
      }
    } catch (err) {
      // キャッシュ障害時は通常処理にフォールバック（動作は従来と同一）
    }
    return getAllRecords();
  } else if (action === 'getSettings') {
    const settings = getSettings(ss);
    return createJsonResponse({ status: 'ok', settings: settings });
  } else if (action === 'ping') {
    return createJsonResponse({ status: 'ok', message: 'pong', timestamp: new Date().toISOString() });
  }

  return createJsonResponse({ status: 'error', message: 'Invalid action' });
}

function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let postData;
  try {
    postData = JSON.parse(e.postData.contents);
  } catch (err) {
    return createJsonResponse({ status: 'error', message: 'Invalid JSON' });
  }

  const action = postData.action;

  // ===== 排他制御（ロック機能）の追加 =====
  // 複数のPCからの通信が重なった際、最大10秒間待機して順番に確実に処理する
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (err) {
    return createJsonResponse({ status: 'error', message: 'Server busy. Please try again.' });
  }

  try {
    if (action === 'record') {
      writeRecord(ss, postData);
      invalidateRecordsCache();
      return createJsonResponse({ status: 'ok' });
    } else if (action === 'sync') {
      const records = postData.records || [];
      records.forEach(r => writeRecord(ss, r));
      invalidateRecordsCache();
      return createJsonResponse({ status: 'ok', count: records.length });
    } else if (action === 'delete') {
      deleteRecord(ss, postData.id);
      invalidateRecordsCache();
      return createJsonResponse({ status: 'ok' });
    } else if (action === 'saveSettings') {
      const result = saveSettings(ss, postData.settings);
      return createJsonResponse(result);
    }
  } finally {
    lock.releaseLock();
  }

  return createJsonResponse({ status: 'error', message: 'Invalid action' });
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===== 全レコード読み取り =====
function getAllRecords() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const tz = Session.getScriptTimeZone();

  // IDをキーとしたMapを使い、同一IDで複数行ある場合は最新を残す
  // clientUpdatedAt (11列目) または serverUpdatedAtMs (10列目) を優先
  const recordMap = {};

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    // シート名は「スタッフ名_YYYYMM」形式を期待
    if (!sheetName.match(/^.+_\d{6}$/)) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    // 12列まで取得（日付、曜日、出勤、中抜け、退勤、賄い、有給、備考、更新日時、clientUpdatedAt、削除フラグ）
    // ただし、最大列数(getMaxColumns)を超えて読み込もうとすると境界外エラーになるため、安全な範囲を指定する
    const lastCol = sheet.getLastColumn();
    const maxCol = sheet.getMaxColumns();
    const colsToRead = Math.min(Math.max(lastCol, 1), maxCol, 13);
    if (colsToRead < 1) return;

    let data;
    try {
      data = sheet.getRange(2, 1, lastRow - 1, colsToRead).getValues();
    } catch (e) {
      console.warn('Sheet access error: ' + sheetName, e);
      return;
    }

    data.forEach(row => {
      // 最小限必要な列データ(日付)があるか確認
      if (!row || row.length < 1) return;
      
      const dateCell = row[0];
      if (!dateCell) return;

      let dateStr = '';
      try {
        if (dateCell instanceof Date) {
          dateStr = Utilities.formatDate(dateCell, tz, 'yyyy-MM-dd');
        } else if (typeof dateCell === 'string') {
          const match = dateCell.match(/(\d{4}-\d{2}-\d{2})/);
          if (!match) return;
          dateStr = match[1];
        } else {
          return;
        }
      } catch (e) {
        return; // 日付パース失敗時はスキップ
      }

      const staffName = sheetName.replace(/_\d{6}$/, '');
      const id = staffName + '_' + dateStr;
      
      const clientUpdatedAt = (row.length > 10 && row[10]) ? String(row[10]) : '';
      const clockIn    = formatTimeCell(row[2]);
      const clockOut   = formatTimeCell(row[5]);
      const breakStart = formatTimeCell(row[3]);
      const breakEnd   = formatTimeCell(row[4]);
      const meal       = row[6] === '有' || row[6] === 1 || row[6] === true;
      const isPaidLeave = (row[7] === '有給') || false;
      const remarks    = (row.length > 8 && row[8]) ? String(row[8]) : '';
      // 列13（休憩分数）: 空セルはnullを返す（旧データと明示的な0を区別するため）
      let additionalBreakMins = null;
      if (row.length > 12 && row[12] !== '' && row[12] !== null && row[12] !== undefined) {
        const parsed = parseInt(row[12]);
        additionalBreakMins = isNaN(parsed) ? null : parsed;
      }

      const isActuallyPaidLeave = isPaidLeave ||
        (typeof remarks === 'string' && remarks.includes('有給申請'));

      // 重複比較用: clientUpdatedAt > サーバー側updatedAt の優先順
      let compareMs = 0;
      let serverUpdatedAtMs = 0;
      try {
        if (row.length > 9 && row[9] instanceof Date) {
          serverUpdatedAtMs = row[9].getTime();
        }
        
        if (clientUpdatedAt && !isNaN(new Date(clientUpdatedAt).getTime())) {
          compareMs = new Date(clientUpdatedAt).getTime();
        } else {
          compareMs = serverUpdatedAtMs;
        }
      } catch (e) {
        compareMs = 0;
      }

      // 12列目（index 11）= deletedフラグ
      let isDeleted = false;
      if (row.length > 11 && (row[11] === 'true' || row[11] === true)) {
        isDeleted = true;
      }

      const record = {
        id: id,
        staffName: staffName,
        date: dateStr,
        clockIn: clockIn || null,
        clockOut: clockOut || null,
        breakStart: breakStart || null,
        breakEnd: breakEnd || null,
        meal: meal,
        isPaidLeave: isActuallyPaidLeave,
        remarks: remarks,
        additionalBreakMins: additionalBreakMins,
        clientUpdatedAt: clientUpdatedAt || null,
        serverUpdatedAtMs: serverUpdatedAtMs,
        deleted: isDeleted
      };

      // 同じIDが複数行ある場合（重複行）は最新のものだけ残す
      if (!recordMap[id] || compareMs > recordMap[id]._compareMs) {
        record._compareMs = compareMs;
        recordMap[id] = record;
      }
    });
  });

  // 比較用フィールドを除いて返す
  const allRecords = Object.values(recordMap).map(r => {
    const rec = Object.assign({}, r);
    delete rec._compareMs;
    return rec;
  });

  const payload = JSON.stringify({
    status: 'ok',
    count: allRecords.length,
    records: allRecords,
    timestamp: new Date().toISOString()
  });

  // 結果を短時間キャッシュ（100KB超過などで保存できない場合は黙ってスキップ＝従来動作）
  try {
    CacheService.getScriptCache().put(RECORDS_CACHE_KEY, payload, RECORDS_CACHE_SECONDS);
  } catch (err) {
    // キャッシュ保存失敗は無視（動作に影響なし）
  }

  return ContentService.createTextOutput(payload)
    .setMimeType(ContentService.MimeType.JSON);
}

// 書き込み発生時にキャッシュを即座に破棄する（古いデータが返るのを防ぐ）
function invalidateRecordsCache() {
  try {
    CacheService.getScriptCache().remove(RECORDS_CACHE_KEY);
  } catch (err) {
    // 破棄失敗時も最大15秒で自然消滅するため致命的ではない
  }
}

// ===== 時刻セルを "HH:MM:SS" 文字列に変換 =====
function formatTimeCell(cell) {
  if (!cell) return null;
  if (cell instanceof Date) {
    return Utilities.formatDate(cell, Session.getScriptTimeZone(), 'HH:mm:ss');
  }
  if (typeof cell === 'number' && cell > 0 && cell < 1) {
    // Excelの時刻小数値 → 時刻文字列
    const totalSeconds = Math.round(cell * 86400);
    const h = Math.floor(totalSeconds / 3600);
    const m = Math.floor((totalSeconds % 3600) / 60);
    const s = totalSeconds % 60;
    return String(h).padStart(2,'0') + ':' + String(m).padStart(2,'0') + ':' + String(s).padStart(2,'0');
  }
  if (typeof cell === 'string' && cell.match(/^\d{1,2}:\d{2}/)) {
    return cell.length === 5 ? cell + ':00' : cell;
  }
  return null;
}

// ===== レコード削除 =====
function deleteRecord(ss, recordId) {
  if (!recordId) return;
  // recordId の形式: "スタッフ名_YYYY-MM-DD"
  const parts = recordId.match(/^(.+)_(\d{4}-\d{2}-\d{2})$/);
  if (!parts) return;

  const staffName = parts[1];
  const dateStr   = parts[2];
  const dateParts = dateStr.split('-');
  const yearMonth = dateParts[0] + dateParts[1];
  const sheetName = staffName + '_' + yearMonth;

  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const tz = Session.getScriptTimeZone();
  const dates = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  // 重複データが存在していても確実にすべて消すため、下（最後）から逆順でループして削除する
  for (let i = dates.length - 1; i >= 0; i--) {
    const cellDate = dates[i][0];
    let cellDateStr = '';
    if (cellDate instanceof Date) {
      cellDateStr = Utilities.formatDate(cellDate, tz, 'yyyy-MM-dd');
    } else if (typeof cellDate === 'string') {
      const match = cellDate.match(/(\d{4}-\d{2}-\d{2})/);
      cellDateStr = match ? match[1] : cellDate;
    }
    if (cellDateStr === dateStr) {
      sheet.deleteRow(i + 2);
    }
  }
}

// ===== データ書き込み =====
function writeRecord(ss, record) {
  const staffName = record.staffName;
  if (!staffName) return;

  const dateStr = record.date; // "2026-03-01" 形式
  if (!dateStr) return;

  // 日付から年月を取得してシート名を決定（例: 草野_202603）
  const dateParts = dateStr.split('-');
  const yearMonth = dateParts[0] + dateParts[1];
  const sheetName = staffName + '_' + yearMonth;

  // シートを取得（なければ作成）
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, 13).setValues([[
      '日付', '曜日', '出勤', '中抜け開始', '中抜け終了', '退勤',
      '賄い', '有給', '備考', '更新日時', 'clientUpdatedAt', '削除フラグ', '休憩分数'
    ]]);
    sheet.getRange(1, 1, 1, 13).setFontWeight('bold').setBackground('#e8f5e9');
    sheet.setFrozenRows(1);
    sheet.setTabColor('#4caf50');
  } else {
    // 既存のシートがあり、最大列数が13列未満の場合は自動的に13列以上に拡張する
    const currentMaxCols = sheet.getMaxColumns();
    if (currentMaxCols < 13) {
      sheet.insertColumnsAfter(currentMaxCols, 13 - currentMaxCols);
      // 足りないヘッダーを自動補完
      const headerRange = sheet.getRange(1, currentMaxCols + 1, 1, 13 - currentMaxCols);
      const allHeaders = [
        '日付', '曜日', '出勤', '中抜け開始', '中抜け終了', '退勤',
        '賄い', '有給', '備考', '更新日時', 'clientUpdatedAt', '削除フラグ', '休憩分数'
      ];
      const missingHeaders = allHeaders.slice(currentMaxCols);
      headerRange.setValues([missingHeaders]);
      headerRange.setFontWeight('bold').setBackground('#e8f5e9');
    }
  }

  // スクリプトのタイムゾーンを使って日付比較（GASデプロイ環境に合わせる）
  const tz = Session.getScriptTimeZone();

  // 受信データの clientUpdatedAt
  const incomingClientUpdatedAt = record.clientUpdatedAt || '';
  const incomingMs = incomingClientUpdatedAt ? new Date(incomingClientUpdatedAt).getTime() || 0 : 0;

  // 既存の行を検索（日付が一致する行を探す）
  const lastRow = sheet.getLastRow();
  let targetRow = -1;
  let existingLatestMs = -1; // 既存行のclientUpdatedAt最大値

  if (lastRow >= 2) {
    const lastCol = sheet.getLastColumn();
    const maxCol = sheet.getMaxColumns();
    const colsToRead = Math.min(Math.max(lastCol, 1), maxCol, 13);
    let allData = [];
    if (colsToRead >= 1) {
      try {
        allData = sheet.getRange(2, 1, lastRow - 1, colsToRead).getValues();
      } catch (e) {
        console.warn('Find existing row error in writeRecord: ' + sheetName, e);
      }
    }
    for (let i = 0; i < allData.length; i++) {
      const cellDate = allData[i][0];
      let cellDateStr = '';
      if (cellDate instanceof Date) {
        cellDateStr = Utilities.formatDate(cellDate, tz, 'yyyy-MM-dd');
      } else if (typeof cellDate === 'string') {
        const match = cellDate.match(/(\d{4}-\d{2}-\d{2})/);
        cellDateStr = match ? match[1] : cellDate;
      }
      if (cellDateStr === dateStr) {
        // 既存行のclientUpdatedAtを取得して比較
        let existingCuAt = '';
        if (allData[i].length > 10 && allData[i][10]) {
          existingCuAt = String(allData[i][10]);
        }
        let existingMs = 0;
        if (existingCuAt) {
          existingMs = new Date(existingCuAt).getTime() || 0;
        } else if (allData[i].length > 9 && allData[i][9] instanceof Date) {
          existingMs = allData[i][9].getTime();
        }

        if (targetRow === -1 || existingMs > existingLatestMs) {
          targetRow = i + 2;
          existingLatestMs = existingMs;
        }
      }
    }
  }

  // Idempotency保護: 受信データが既存データより古い場合は書き込みをスキップ
  if (targetRow !== -1 && incomingMs > 0 && existingLatestMs > 0 && incomingMs < existingLatestMs) {
    return; // 古いデータなのでスキップ
  }

  if (targetRow === -1) {
    targetRow = lastRow + 1;
  }
  
  if (targetRow < 2) targetRow = 2;

  // 曜日を計算
  const dateObj = new Date(parseInt(dateParts[0]), parseInt(dateParts[1]) - 1, parseInt(dateParts[2]));
  const dowNames = ['日', '月', '火', '水', '木', '金', '土'];
  const dow = dowNames[dateObj.getDay()];

  // フロントエンドは常にレコードの全フィールドを送信するため、
  // 受け取ったデータをそのまま書き込む（13列に拡張: 休憩分数）
  const rowData = [
    dateStr,
    dow,
    record.clockIn || '',
    record.breakStart || '',
    record.breakEnd || '',
    record.clockOut || '',
    record.meal ? '有' : '',
    record.isPaidLeave ? '有給' : '',
    record.remarks || '',
    new Date(),
    incomingClientUpdatedAt,
    record.deleted ? 'true' : '',
    (record.additionalBreakMins !== undefined && record.additionalBreakMins !== null) ? record.additionalBreakMins : 0
  ];

  sheet.getRange(targetRow, 1, 1, 13).setValues([rowData]);
}

// ===== 設定情報の操作 =====
const SETTINGS_SHEET_NAME = 'AppSettings';

function getSettings(ss) {
  let sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!sheet || sheet.getLastRow() === 0) {
    return null; // まだ設定がない場合
  }
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(1, 1, lastRow, 2).getValues();
  let settings = {};
  data.forEach(row => {
    if (row[0] === 'staffList') {
      try {
        settings.staffList = JSON.parse(row[1]);
      } catch(e) { settings.staffList = []; }
    }
    if (row[0] === 'adminPin') {
      // 数値に変換されていた場合も考慮して文字列化し、4桁未満なら0埋めする
      const pin = row[1].toString();
      settings.adminPin = pin.length < 4 && /^\d+$/.test(pin) ? pin.padStart(4, '0') : pin;
    }
  });
  return settings;
}

// フロントエンドのローカル初期値名簿（仮データ）。この名簿での上書き要求はサーバー側でも拒否する。
// （障害⑥: 初期値キャッシュの端末から本番AppSettingsが上書きされた事故の再発防止・二重防御）
const DEFAULT_SEED_NAMES = ["草野", "田中", "小出", "高橋", "小田", "林", "山田", "市川（悠）", "谷口", "森川", "五十嵐", "市川（公）"];

function isDefaultSeedList(list) {
  if (!Array.isArray(list) || list.length !== DEFAULT_SEED_NAMES.length) return false;
  const names = list.map(function(s) {
    return (typeof s === 'string') ? s : ((s && s.name) || '');
  }).sort();
  const seed = DEFAULT_SEED_NAMES.slice().sort();
  return names.every(function(n, i) { return n === seed[i]; });
}

// 受信した設定のバリデーション。不正なら理由文字列を返し、正常なら null を返す
function validateIncomingSettings(ss, settings) {
  if (!settings || typeof settings !== 'object') {
    return '設定オブジェクトが不正です';
  }
  if (settings.staffList !== undefined) {
    if (!Array.isArray(settings.staffList)) {
      return 'staffListが配列ではありません';
    }
    if (settings.staffList.length === 0) {
      return 'staffListが空です（全削除は手動でのみ許可）';
    }
    const hasInvalid = settings.staffList.some(function(s) {
      const name = (typeof s === 'string') ? s : ((s && s.name) || '');
      return !name || typeof name !== 'string' || !name.trim();
    });
    if (hasInvalid) {
      return 'staffListに名前が欠損した要素が含まれています';
    }
    if (isDefaultSeedList(settings.staffList)) {
      return '初期値名簿（仮データ）による上書きは禁止されています';
    }
    // 大幅減少ガード: 現名簿10名以上に対し、半数未満への縮小要求は誤送信とみなして拒否
    const current = getSettings(ss);
    if (current && Array.isArray(current.staffList) && current.staffList.length >= 10) {
      if (settings.staffList.length < current.staffList.length / 2) {
        return '現在の名簿(' + current.staffList.length + '名)から' + settings.staffList.length +
               '名への大幅減少要求のため拒否しました（誤送信保護）';
      }
    }
  }
  return null;
}

function saveSettings(ss, settings) {
  // ===== サーバー側バリデーション（本番データ保護の最終防衛線） =====
  const validationError = validateIncomingSettings(ss, settings);
  if (validationError) {
    console.warn('saveSettings rejected: ' + validationError);
    return { status: 'rejected', message: validationError };
  }

  // ===== 上書き前の現行設定を履歴シートへ退避（ロールバック用） =====
  try {
    const previous = getSettings(ss);
    if (previous && previous.staffList && previous.staffList.length > 0) {
      appendSettingsHistory(ss, previous, '[上書き前の自動退避]');
    }
  } catch(e) {
    console.error('Pre-save backup failed:', e);
  }

  let sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SETTINGS_SHEET_NAME);
    sheet.setTabColor('#ff9800');
  }
  sheet.clear();
  const rows = [];
  if (settings.staffList) rows.push(['staffList', JSON.stringify(settings.staffList)]);
  if (settings.adminPin !== undefined && settings.adminPin !== null) rows.push(['adminPin', settings.adminPin]);
  rows.push(['updatedAt', new Date()]);

  if (rows.length > 0) {
    const range = sheet.getRange(1, 1, rows.length, 2);
    // B列をテキスト形式に設定してスプレッドシートの自動数値変換を防ぐ
    sheet.getRange(1, 2, rows.length, 1).setNumberFormat('@');
    range.setValues(rows);
  }

  // ===== 設定変更履歴シートへの自動記録 =====
  try {
    appendSettingsHistory(ss, settings, '');
  } catch(e) {
    console.error('Settings backup history failed:', e);
  }

  return { status: 'ok' };
}

// 履歴シート（AppSettingsHistory）への追記共通処理
function appendSettingsHistory(ss, settings, note) {
  const HISTORY_SHEET_NAME = 'AppSettingsHistory';
  let historySheet = ss.getSheetByName(HISTORY_SHEET_NAME);
  if (!historySheet) {
    historySheet = ss.insertSheet(HISTORY_SHEET_NAME);
    historySheet.setTabColor('#4caf50');
    historySheet.appendRow(['バックアップ日時', 'スタッフ人数', 'スタッフ一覧JSON', '管理者PIN']);
    historySheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#e8f5e9');
  }
  const count = settings.staffList ? settings.staffList.length : 0;
  const jsonStr = settings.staffList ? JSON.stringify(settings.staffList) : '';
  historySheet.appendRow([
    Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss') + (note ? ' ' + note : ''),
    count + '名',
    jsonStr,
    settings.adminPin || ''
  ]);
}

/**
 * 毎日自動で現在の設定を履歴バックアップシートに記録する定期実行用関数
 * ※ GASの「トリガー」画面から、この関数を毎日（例: 深夜〜早朝）実行するよう設定可能です。
 */
function recordDailySettingsBackup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = getSettings(ss);
  if (!settings || !settings.staffList) return;
  
  appendSettingsHistory(ss, settings, '[毎日自動バックアップ]');
}
