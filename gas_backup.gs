/**
 * タイムカード サーバーサイド(Google Apps Script)
 * スプレッドシートをDBとして使用します。
 */

// ===== 全件取得キャッシュ =====
// 全シート読み込み（5〜6秒）の結果を短時間だけサーバー側に記憶して応答を高速化する。
// 打刻・削除などの書き込みがあった瞬間に破棄するため、古いデータが返り続けることはない。
const RECORDS_CACHE_KEY = 'getAll_cache_v1';
// 60秒キャッシュ。書き込み時は即破棄するため打刻の反映は遅れない。
// （スプレッドシートを直接手編集した場合のみ最大60秒反映が遅れる）
const RECORDS_CACHE_SECONDS = 60;

function doGet(e) {
  const action = e.parameter.action;
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (action === 'getAll') {
    let cached = getRecordsCache();
    if (cached) {
      return ContentService.createTextOutput(cached)
        .setMimeType(ContentService.MimeType.JSON);
    }
    // ===== キャッシュ再構築の同時多発（スタンピード）防止 =====
    // キャッシュ切れの瞬間に複数端末から重い全件読み込みが同時実行されると
    // GASの同時実行上限に達して404が返るため、再構築は1つに直列化する。
    // ※ doPost（打刻）はgetScriptLockを使用しており、こちらはgetDocumentLockなので
    //    打刻処理をブロックすることはない。
    const lock = LockService.getDocumentLock();
    let locked = false;
    try { locked = lock.tryLock(25000); } catch (err) { locked = false; }
    if (locked) {
      try {
        // ロック待ちの間に他の実行がキャッシュを作った可能性を再確認
        cached = getRecordsCache();
        if (cached) {
          return ContentService.createTextOutput(cached)
            .setMimeType(ContentService.MimeType.JSON);
        }
        return getAllRecords();
      } finally {
        lock.releaseLock();
      }
    }
    // ロックが取れなかった場合も従来通り自力で構築（フォールバック）
    return getAllRecords();
  } else if (action === 'getSettings') {
    const settings = getSettings(ss);
    return createJsonResponse({ status: 'ok', settings: settings });
  } else if (action === 'ping') {
    return createJsonResponse({ status: 'ok', message: 'pong', timestamp: new Date().toISOString() });
  } else if (action === 'getLeaveInfo') {
    const year = e.parameter.year || String(new Date().getFullYear());
    const result = getLeaveInfo(year);
    return createJsonResponse({ status: 'ok', found: result.found, fileName: result.fileName, leaveInfo: result.leaveInfo });
  }

  return createJsonResponse({ status: 'error', message: 'Invalid action' });
}

// ===== 有給休暇管理簿（別スプレッドシート・オーナー管轄）から残日数・次回付与日を読み取る =====
// このシートはオーナーが手作業で更新している「正本」。ここでは読み取るのみで書き込みは行わない。
// ファイルは年ごとに新規作成される（例:「年次有給休暇管理簿2026」→翌年は「年次有給休暇管理簿2027」）ため、
// IDを直接指定せず、決まったフォルダ内をファイル名（年）で毎回検索する。
const LEAVE_FOLDER_ID = '1ltoDDydrHBW1Yu6R3FX04ffd8Sa-0QEr'; // Googleドライブの「有給」フォルダ
const LEAVE_FILE_NAME_PREFIX = '年次有給休暇管理簿';
const LEAVE_SHEET_NAME = '管理簿';
const LEAVE_NAME_COL = 2;      // B列: 名前
const LEAVE_NEXT_GRANT_COL = 9; // I列: 基準日（次回付与日）
const LEAVE_REMAINING_COL = 14; // N列: 有給残日数

function getLeaveInfo(year) {
  const result = {};
  try {
    const folder = DriveApp.getFolderById(LEAVE_FOLDER_ID);
    const keyword = LEAVE_FILE_NAME_PREFIX + year;
    const files = folder.searchFiles(`title contains '${keyword.replace(/'/g, "\\'")}' and mimeType = 'application/vnd.google-apps.spreadsheet'`);
    if (!files.hasNext()) {
      // 見つからない場合はアプリ側でアラート表示できるよう found:false を返す（打刻機能自体は止めない）
      return { found: false, fileName: null, leaveInfo: result };
    }
    const file = files.next();
    const leaveSs = SpreadsheetApp.open(file);
    const sheet = leaveSs.getSheetByName(LEAVE_SHEET_NAME);
    if (!sheet) return { found: false, fileName: file.getName(), leaveInfo: result };
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { found: true, fileName: file.getName(), leaveInfo: result };
    const values = sheet.getRange(2, 1, lastRow - 1, LEAVE_REMAINING_COL).getValues();
    for (const row of values) {
      const name = String(row[LEAVE_NAME_COL - 1] || '').trim();
      if (!name) continue;
      const nextGrantRaw = row[LEAVE_NEXT_GRANT_COL - 1];
      const nextGrant = nextGrantRaw instanceof Date
        ? Utilities.formatDate(nextGrantRaw, Session.getScriptTimeZone(), 'yyyy/MM/dd')
        : (nextGrantRaw || '');
      result[name] = {
        remainingDays: row[LEAVE_REMAINING_COL - 1],
        nextGrantDate: nextGrant,
      };
    }
    return { found: true, fileName: file.getName(), leaveInfo: result };
  } catch (e) {
    // 有給管理表が読めない場合も打刻機能自体は止めない（見つからなかった扱いにする）
    console.error('getLeaveInfo error: ' + e.message);
    return { found: false, fileName: null, leaveInfo: result };
  }
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
    } else if (action === 'confirmMonth') {
      // スタッフ本人による「今月分を確認しました」の記録（押すたびに日時を上書き）
      const result = confirmMonth(ss, postData.staffName, postData.yearMonth);
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
    const colsToRead = Math.min(Math.max(lastCol, 1), maxCol, 14);
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

      // 列14（交通費）: 空セルはnull
      let transport = null;
      if (row.length > 13 && row[13] !== '' && row[13] !== null && row[13] !== undefined) {
        const t = parseInt(row[13]);
        transport = isNaN(t) ? null : t;
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
        transport: transport,
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

  // 結果を短時間キャッシュ（保存できない場合は黙ってスキップ＝従来動作）
  putRecordsCache(payload);

  return ContentService.createTextOutput(payload)
    .setMimeType(ContentService.MimeType.JSON);
}

// ===== キャッシュ分割保存 =====
// CacheServiceは1キーあたり100KBまでしか保存できないため、
// 応答JSON（現在135KB超）を90KBずつのチャンクに分割して保存・復元する。
function putRecordsCache(payload) {
  try {
    const cache = CacheService.getScriptCache();
    const CHUNK_SIZE = 90000;
    const entries = {};
    let n = 0;
    for (let i = 0; i < payload.length; i += CHUNK_SIZE) {
      entries[RECORDS_CACHE_KEY + '_' + n] = payload.substring(i, i + CHUNK_SIZE);
      n++;
    }
    entries[RECORDS_CACHE_KEY + '_meta'] = String(n);
    cache.putAll(entries, RECORDS_CACHE_SECONDS);
  } catch (err) {
    // キャッシュ保存失敗は無視（動作に影響なし）
  }
}

function getRecordsCache() {
  try {
    const cache = CacheService.getScriptCache();
    const meta = cache.get(RECORDS_CACHE_KEY + '_meta');
    if (!meta) return null;
    const n = parseInt(meta);
    if (!n || n < 1) return null;
    const keys = [];
    for (let i = 0; i < n; i++) keys.push(RECORDS_CACHE_KEY + '_' + i);
    const parts = cache.getAll(keys);
    let out = '';
    for (let i = 0; i < n; i++) {
      const p = parts[RECORDS_CACHE_KEY + '_' + i];
      if (p === undefined || p === null) return null; // チャンク欠損時はキャッシュ不成立として通常処理へ
      out += p;
    }
    return out;
  } catch (err) {
    return null;
  }
}

// 書き込み発生時にキャッシュを即座に破棄する（古いデータが返るのを防ぐ）
// metaキーを消せばチャンクへの参照が失われ、キャッシュ全体が無効になる
function invalidateRecordsCache() {
  try {
    CacheService.getScriptCache().remove(RECORDS_CACHE_KEY + '_meta');
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
    sheet.getRange(1, 1, 1, 14).setValues([[
      '日付', '曜日', '出勤', '中抜け開始', '中抜け終了', '退勤',
      '賄い', '有給', '備考', '更新日時', 'clientUpdatedAt', '削除フラグ', '休憩分数', '交通費'
    ]]);
    sheet.getRange(1, 1, 1, 14).setFontWeight('bold').setBackground('#e8f5e9');
    sheet.setFrozenRows(1);
    sheet.setTabColor('#4caf50');
  } else {
    // 既存のシートがあり、最大列数が13列未満の場合は自動的に13列以上に拡張する
    const currentMaxCols = sheet.getMaxColumns();
    if (currentMaxCols < 14) {
      sheet.insertColumnsAfter(currentMaxCols, 14 - currentMaxCols);
      // 足りないヘッダーを自動補完
      const headerRange = sheet.getRange(1, currentMaxCols + 1, 1, 14 - currentMaxCols);
      const allHeaders = [
        '日付', '曜日', '出勤', '中抜け開始', '中抜け終了', '退勤',
        '賄い', '有給', '備考', '更新日時', 'clientUpdatedAt', '削除フラグ', '休憩分数', '交通費'
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
    const colsToRead = Math.min(Math.max(lastCol, 1), maxCol, 14);
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
  // 受け取ったデータをそのまま書き込む（14列: 休憩分数・交通費）
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
    (record.additionalBreakMins !== undefined && record.additionalBreakMins !== null) ? record.additionalBreakMins : 0,
    (record.transport !== undefined && record.transport !== null && record.transport !== '') ? Number(record.transport) : ''
  ];

  // ===== 修正履歴（EditLog）: 既存行の内容が変わった場合のみ自動で1行追記 =====
  try {
    if (targetRow <= lastRow && lastRow >= 2) {
      logEditIfChanged(ss, sheetName, sheet, targetRow, rowData);
    }
  } catch (e) {
    console.warn('EditLog failed:', e);
  }

  sheet.getRange(targetRow, 1, 1, 14).setValues([rowData]);
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
  // 月次確認（スタッフが「今月分を確認しました」を押した日時）を同梱
  try {
    const conf = ss.getSheetByName('Confirmations');
    if (conf && conf.getLastRow() >= 2) {
      const rows = conf.getRange(2, 1, conf.getLastRow() - 1, 3).getValues();
      const map = {};
      rows.forEach(function(r) {
        if (r[0] && r[1]) map[String(r[0]) + '|' + String(r[1])] = String(r[2] || '');
      });
      settings.confirmations = map;
    }
  } catch (e) { /* 確認シートが無い場合は無視 */ }
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


// ===== 月次確認の記録（押すたびに日時を上書き） =====
function confirmMonth(ss, staffName, yearMonth) {
  if (!staffName || !/^\d{4}-\d{2}$/.test(String(yearMonth || ''))) {
    return { status: 'error', message: 'スタッフ名と対象月を指定してください' };
  }
  let sheet = ss.getSheetByName('Confirmations');
  if (!sheet) {
    sheet = ss.insertSheet('Confirmations');
    sheet.setTabColor('#3f51b5');
    sheet.appendRow(['スタッフ名', '対象月', '確認日時']);
    sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#e8eaf6');
  }
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const rows = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) === staffName && String(rows[i][1]) === yearMonth) {
        sheet.getRange(i + 2, 3).setValue(now); // 上書き（最後に押した日時が有効）
        return { status: 'ok', confirmedAt: now };
      }
    }
  }
  sheet.appendRow([staffName, yearMonth, now]);
  return { status: 'ok', confirmedAt: now };
}

// ===== 修正履歴（EditLog）: 既存行の値が変わったときだけ自動追記 =====
function logEditIfChanged(ss, sheetName, sheet, targetRow, newRowData) {
  const old = sheet.getRange(targetRow, 1, 1, 14).getValues()[0];
  const tz = Session.getScriptTimeZone();
  const fmt = function(v) {
    if (v === null || v === undefined || v === '') return '';
    if (v instanceof Date) return Utilities.formatDate(v, tz, 'HH:mm');
    return String(v);
  };
  // 比較対象: 出勤(3) 中抜け開始(4) 中抜け終了(5) 退勤(6) 賄い(7) 有給(8) 備考(9) 休憩分数(13) 交通費(14)
  const cols = [
    [2, '出勤'], [3, '中抜け開始'], [4, '中抜け終了'], [5, '退勤'],
    [6, '賄い'], [7, '有給'], [8, '備考'], [12, '休憩分数'], [13, '交通費']
  ];
  const changes = [];
  cols.forEach(function(c) {
    const o = fmt(old[c[0]]);
    const n = fmt(newRowData[c[0]]);
    if (o !== n) changes.push(c[1] + ': 「' + (o || 'なし') + '」→「' + (n || 'なし') + '」');
  });
  if (changes.length === 0) return;

  let log = ss.getSheetByName('EditLog');
  if (!log) {
    log = ss.insertSheet('EditLog');
    log.setTabColor('#9e9e9e');
    log.appendRow(['修正日時', 'スタッフ', '対象日', '変更内容']);
    log.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#eeeeee');
  }
  log.appendRow([
    Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss'),
    sheetName.replace(/_\d{6}$/, ''),
    fmtDateCell(old[0], tz),
    changes.join(' / ')
  ]);
}

function fmtDateCell(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  return String(v || '');
}
