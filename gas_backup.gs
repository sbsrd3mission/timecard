/**
 * タイムカード 双方向同期 - Google Apps Script
 * 
 * このスクリプトをGoogleスプレッドシートのApps Scriptにコピーして
 * Webアプリとしてデプロイしてください。
 * 
 * デプロイ設定:
 *   - 実行するユーザー: 自分
 *   - アクセスできるユーザー: 全員（匿名ユーザーを含む）
 * 
 * 機能:
 * [書き込み] POST: 打刻データを受信してスプレッドシートに記録
 * [読み取り] GET ?action=getAll: 全打刻データをJSON形式で返す
 * [読み取り] GET ?action=ping: 接続確認
 */

// ===== CORS ヘッダーを付けたレスポンス生成 =====
function createJsonResponse(data) {
  const output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ===== GET: データ読み取りエンドポイント =====
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'ping';

    if (action === 'getAll') {
      return getAllRecords();
    }

    if (action === 'getSettings') {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      return createJsonResponse({ status: 'ok', settings: getSettings(ss) });
    }

    // ping（接続テスト）
    return createJsonResponse({
      status: 'ok',
      message: '接続成功！タイムカード双方向同期が動作中です。',
      timestamp: new Date().toISOString()
    });

  } catch (error) {
    return createJsonResponse({ status: 'error', message: error.message });
  }
}

// ===== POST: データ書き込みエンドポイント =====
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (data.action === 'sync') {
      // 複数レコードの一括同期
      const records = data.records || [];
      let count = 0;
      records.forEach(record => {
        writeRecord(ss, record);
        count++;
      });
      return createJsonResponse({
        status: 'ok',
        message: count + '件のデータを同期しました',
        count: count
      });
    }

    if (data.action === 'record') {
      // 単一レコードの書き込み
      writeRecord(ss, data);
      return createJsonResponse({ status: 'ok', message: 'データを記録しました' });
    }

    if (data.action === 'delete') {
      // レコードの削除（論理削除: Tombstone）
      // 古いDELETE方式から、deletedフラグを立てて保存する方式に変更
      data.deleted = true;
      if (!data.clientUpdatedAt) data.clientUpdatedAt = new Date().toISOString();
      writeRecord(ss, data);
      return createJsonResponse({ status: 'ok', message: 'データを削除しました（論理削除）' });
    }

    if (data.action === 'saveSettings') {
      // 設定情報の保存（スタッフリスト、PINコード）
      saveSettings(ss, data.settings);
      return createJsonResponse({ status: 'ok', message: '設定を保存しました' });
    }

    return createJsonResponse({ status: 'error', message: '不明なアクションです' });

  } catch (error) {
    return createJsonResponse({ status: 'error', message: error.message });
  }
}

// ===== 全レコード読み取り =====
function getAllRecords() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const tz = Session.getScriptTimeZone(); // スクリプトのタイムゾーン（通常 Asia/Tokyo）

  // IDをキーとしたMapを使い、同一IDで複数行ある場合は最新を残す
  // clientUpdatedAt（11列目）を優先し、なければサーバー側updatedAt（10列目）で比較
  const recordMap = {};

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    // シート名は「スタッフ名_YYYYMM」形式を期待
    if (!sheetName.match(/^.+_\d{6}$/)) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    // ヘッダー行を含む全データを取得（11列に拡張）
    const lastCol = Math.max(sheet.getLastColumn(), 12);
    const data = sheet.getRange(2, 1, lastRow - 1, Math.min(lastCol, 12)).getValues();

    data.forEach(row => {
      const dateCell = row[0];
      if (!dateCell) return;

      let dateStr = '';
      if (dateCell instanceof Date) {
        dateStr = Utilities.formatDate(dateCell, tz, 'yyyy-MM-dd');
      } else if (typeof dateCell === 'string') {
        const match = dateCell.match(/(\d{4}-\d{2}-\d{2})/);
        if (!match) return;
        dateStr = match[1];
      } else {
        return;
      }

      const staffName = sheetName.replace(/_\d{6}$/, '');
      const id = staffName + '_' + dateStr;

      const clockIn    = formatTimeCell(row[2]);
      const clockOut   = formatTimeCell(row[5]);
      const breakStart = formatTimeCell(row[3]);
      const breakEnd   = formatTimeCell(row[4]);
      const meal       = row[6] === '有' || row[6] === 1;
      const isPaidLeave = (row[7] === '有給') || false;
      const remarks    = row[8] || '';

      const isActuallyPaidLeave = isPaidLeave ||
        (typeof remarks === 'string' && remarks.includes('有給申請'));

      // 11列目（index 10）= clientUpdatedAt（ISO文字列）
      // 10列目（index 9）= サーバー側更新日時（フォールバック用）
      let clientUpdatedAt = '';
      if (row.length > 10 && row[10]) {
        clientUpdatedAt = String(row[10]);
      }

      // 重複比較用: clientUpdatedAt > サーバー側updatedAt の優先順
      let compareMs = 0;
      if (clientUpdatedAt) {
        compareMs = new Date(clientUpdatedAt).getTime() || 0;
      } else if (row[9] instanceof Date) {
        compareMs = row[9].getTime();
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
        additionalBreakMins: 0,
        clientUpdatedAt: clientUpdatedAt || null,
        serverUpdatedAtMs: (row[9] instanceof Date) ? row[9].getTime() : 0,
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

  return createJsonResponse({
    status: 'ok',
    count: allRecords.length,
    records: allRecords,
    timestamp: new Date().toISOString()
  });
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
    sheet.getRange(1, 1, 1, 12).setValues([[
      '日付', '曜日', '出勤', '中抜け開始', '中抜け終了', '退勤',
      '賄い', '有給', '備考', '更新日時', 'clientUpdatedAt', '削除フラグ'
    ]]);
    sheet.getRange(1, 1, 1, 12).setFontWeight('bold').setBackground('#e8f5e9');
    sheet.setFrozenRows(1);
    sheet.setTabColor('#4caf50');
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
  let existingLatestClientUpdatedAt = '';

  if (lastRow >= 2) {
    const lastCol = Math.max(sheet.getLastColumn(), 12);
    const allData = sheet.getRange(2, 1, lastRow - 1, Math.min(lastCol, 12)).getValues();
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
        } else if (allData[i][9] instanceof Date) {
          existingMs = allData[i][9].getTime();
        }

        if (targetRow === -1 || existingMs > existingLatestMs) {
          targetRow = i + 2;
          existingLatestMs = existingMs;
          existingLatestClientUpdatedAt = existingCuAt;
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

  // 曜日を計算
  const dateObj = new Date(parseInt(dateParts[0]), parseInt(dateParts[1]) - 1, parseInt(dateParts[2]));
  const dowNames = ['日', '月', '火', '水', '木', '金', '土'];
  const dow = dowNames[dateObj.getDay()];

  // フロントエンドは常にレコードの全フィールドを送信するため、
  // 受け取ったデータをそのまま書き込む（12列に拡張: 削除フラグ）
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
    record.deleted ? 'true' : ''
  ];

  sheet.getRange(targetRow, 1, 1, 12).setValues([rowData]);
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

function saveSettings(ss, settings) {
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
}
