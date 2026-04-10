/**
 * タイムカード自動バックアップ - Google Apps Script
 * 
 * このスクリプトをGoogleスプレッドシートのApps Scriptにコピーして
 * Webアプリとしてデプロイしてください。
 * 
 * 機能:
 * - 打刻データを受信してスプレッドシートに記録
 * - スタッフごと・月ごとにシートを自動作成（例: 草野_202603）
 * - 同じ日のデータは上書き更新
 */

// === 受信処理 ===

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
      return ContentService.createTextOutput(JSON.stringify({
        status: 'ok',
        message: count + '件のデータを同期しました',
        count: count
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    if (data.action === 'record') {
      // 単一レコードの書き込み
      writeRecord(ss, data);
      return ContentService.createTextOutput(JSON.stringify({
        status: 'ok',
        message: 'データを記録しました'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: '不明なアクションです'
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: error.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  // 接続テスト用
  return ContentService.createTextOutput(JSON.stringify({
    status: 'ok',
    message: '接続成功！タイムカードバックアップが動作中です。',
    timestamp: new Date().toISOString()
  })).setMimeType(ContentService.MimeType.JSON);
}

// === データ書き込み ===

function writeRecord(ss, record) {
  const staffName = record.staffName;
  if (!staffName) return;
  
  const dateStr = record.date; // "2026-03-01" 形式
  if (!dateStr) return;
  
  // 日付から年月を取得してシート名を決定（例: 草野_202603）
  const dateParts = dateStr.split('-');
  const yearMonth = dateParts[0] + dateParts[1]; // "202603"
  const yearMonthLabel = dateParts[0] + '年' + parseInt(dateParts[1]) + '月';
  const sheetName = staffName + '_' + yearMonth;
  
  // シートを取得（なければ作成）
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    // ヘッダー行を作成（出勤→中抜け開始→中抜け終了→退勤の順）
    sheet.getRange(1, 1, 1, 10).setValues([[
      '日付', '曜日', '出勤', '中抜け開始', '中抜け終了', '退勤',
      '賄い', '有給', '備考', '更新日時'
    ]]);
    // ヘッダーの書式設定
    sheet.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#e8f5e9');
    sheet.setFrozenRows(1);
    // タイトル的にシート名の色を設定
    sheet.setTabColor('#4caf50');
  }
  
  // 既存の行を検索（日付が一致する行を探す）
  const lastRow = sheet.getLastRow();
  let targetRow = -1;
  
  if (lastRow >= 2) {
    const dates = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < dates.length; i++) {
      const cellDate = dates[i][0];
      let cellDateStr = '';
      if (cellDate instanceof Date) {
        cellDateStr = Utilities.formatDate(cellDate, 'Asia/Tokyo', 'yyyy-MM-dd');
      } else if (typeof cellDate === 'string') {
        cellDateStr = cellDate;
      }
      if (cellDateStr === dateStr) {
        targetRow = i + 2; // 1-indexed + ヘッダー分
        break;
      }
    }
  }
  
  // 新規行の場合は末尾に追加
  if (targetRow === -1) {
    targetRow = lastRow + 1;
  }
  
  // 曜日を計算
  const dateObj = new Date(parseInt(dateParts[0]), parseInt(dateParts[1]) - 1, parseInt(dateParts[2]));
  const dowNames = ['日', '月', '火', '水', '木', '金', '土'];
  const dow = dowNames[dateObj.getDay()];
  
  // データを書き込み（出勤→中抜け開始→中抜け終了→退勤の順）
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
    new Date() // 更新日時
  ];
  
  sheet.getRange(targetRow, 1, 1, 10).setValues([rowData]);
}
