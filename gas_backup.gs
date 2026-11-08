/**
 * タイムカード サーバーサイド(Google Apps Script)
 * スプレッドシートをDBとして使用します。
 *
 * 【2026-11-08 改修（設計書 docs/修正設計書_2026-11-08_… 3.1〜3.5）】
 * - 書き込みは「操作（op）」単位: punch / edit / delete / confirm / legacy。clientVersion 必須。
 * - 行の17列目に「項目ごとの最後の操作時刻(t)・手修正済み印(m)」をJSONで持つ（3.6-2）。
 * - 新旧は端末で操作した時刻（at）で決める（3.6-1）。時計が2分以上ずれた端末は警告する。
 * - 15列目 = 送信元、16列目 = 本人修正の印（3.6-3）。
 * - 処理した opId と結果は OpLog シートに残す（180日分）。同じ opId には最初の結果を返す（3.6-4・3.7-5）。
 * - 操作の時刻は「端末の操作時刻＋（受信時刻－端末の送信時刻）」に直して使う（3.7-1）。最初に受け付けたときの値を OpLog に残す（3.8-2）。
 */

// ===== 版の印（設計書 3.5-5） =====
// この印が無い勤怠の書き込み（古いアプリの record / sync など）は受け付けない。認証ではない。
const MIN_CLIENT_VERSION = 3;

// ===== シートの列（1始まり） =====
const COL = {
  date: 1, dow: 2, clockIn: 3, breakStart: 4, breakEnd: 5, clockOut: 6,
  meal: 7, isPaidLeave: 8, remarks: 9, serverUpdatedAt: 10, clientUpdatedAt: 11,
  deleted: 12, additionalBreakMins: 13, transport: 14, sender: 15, selfEdited: 16, meta: 17
};
const SHEET_COLS = 17;
const SHEET_HEADERS = [
  '日付', '曜日', '出勤', '中抜け開始', '中抜け終了', '退勤',
  '賄い', '有給', '備考', '更新日時', 'clientUpdatedAt', '削除フラグ', '休憩分数', '交通費',
  '送信元', '本人修正', '項目メタ'
];
const TIME_FIELDS = ['clockIn', 'breakStart', 'breakEnd', 'clockOut'];
const ATTR_FIELDS = ['meal', 'transport', 'remarks', 'additionalBreakMins'];
// edit で書ける項目（3.5-2）
const EDIT_FIELDS = TIME_FIELDS.concat(ATTR_FIELDS).concat(['isPaidLeave']);
// 時計のずれの警告（3.6-1）
const CLOCK_SKEW_WARN_MS = 2 * 60 * 1000;
// OpLog（3.6-4）
const OPLOG_SHEET_NAME = 'OpLog';
const OPLOG_KEEP_DAYS = 180;
// OpLog に無く、操作時刻がこれより古い操作は当てない（3.8-8）
const OP_TOO_OLD_DAYS = 180;

// ===== 試験中は本番に書けないようにする（3.8-7） =====
// スクリプトのプロパティに MODE=test と TEST_ALLOWED_FILE_IDS（使ってよいファイルIDをカンマ区切り）を登録すると、
// 登録外のファイルを開こうとしたら処理を止める。本番のファイルIDは下に書いておき、MODE=test では必ず止める。
// 【コードレビュー1-10】本番の ID はコードに書かず、スクリプト プロパティ PROD_TIMECARD_SS_IDS（カンマ区切り）に置く。
// 本番の Apps Script で registerProductionIds() を1回実行すると自動で登録される（反映手順書 5）。
// 下はコードに直接書きたいとき用の予備（ふだんは空のまま）。
// 【2026-11-08 オーナー決定】有給の月末計算をやめたので、有給管理簿の ID の登録（PROD_LEAVE_BOOK_IDS）は使わない
// （管理簿は今までどおり読むだけ）。
const PROD_TIMECARD_SS_IDS = [/* 本番「タイムカードバックアップ」のID */];
function propIdList(key) {
  return String(PropertiesService.getScriptProperties().getProperty(key) || '').split(',').map(function(x) { return x.trim(); }).filter(Boolean);
}
function prodTimecardIds() { return PROD_TIMECARD_SS_IDS.concat(propIdList('PROD_TIMECARD_SS_IDS')); }
// 【3.9-7】実行中のスプレッドシートが本番の ID でなければ、MODE=test が無くても試験扱いにする（設定し忘れの防止）。
// ※ 本番 ID が未登録のときは自動判定できないので、MODE だけで決める。
let TEST_MODE_MEMO = null;
function isTestMode() {
  if (TEST_MODE_MEMO !== null) return TEST_MODE_MEMO;
  let test = PropertiesService.getScriptProperties().getProperty('MODE') === 'test';
  const prodIds = prodTimecardIds();
  if (!test && prodIds.length > 0) {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    test = !active || prodIds.indexOf(active.getId()) === -1;
  }
  TEST_MODE_MEMO = test;
  return test;
}
// 試験モードでは本番の勤怠ファイルは必ず開かない。この GAS が入っている試験用のスプレッドシート（名前に「試験」を含むもの）はそのまま使える。
// それ以外のファイル（戻すときの控えなど）は TEST_ALLOWED_FILE_IDS に書いたものだけ開く。
function assertFileAllowed(id, what) {
  if (!isTestMode()) return;
  if (prodTimecardIds().indexOf(id) !== -1) throw new Error('試験モードでは本番のファイルを開けません: ' + (what || id));
  const active = SpreadsheetApp.getActiveSpreadsheet();
  // この GAS が入っているファイルは、名前に「試験」を含むときだけ許可リストなしで使える（コードレビュー6-1）。
  // 本番の「タイムカードバックアップ」に誤って MODE=test を入れても、ここで止まって本番には書かない。
  if (active && active.getId() === id && String(active.getName()).indexOf('試験') !== -1) return;
  const allowed = propIdList('TEST_ALLOWED_FILE_IDS');
  if (allowed.indexOf(id) === -1) throw new Error('試験モードで許可されていないファイルです: ' + (what || id));
}
function getTimecardSs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  assertFileAllowed(ss.getId(), 'タイムカードのスプレッドシート');
  return ss;
}
function openSpreadsheetChecked(id, what) {
  assertFileAllowed(id, what);
  return SpreadsheetApp.openById(id);
}

// ===== 全件取得キャッシュ =====
// 全シート読み込み（5〜6秒）の結果を短時間だけサーバー側に記憶して応答を高速化する。
// 打刻・削除などの書き込みがあった瞬間に破棄するため、古いデータが返り続けることはない。
// 【2026-11-08】範囲（直近2か月・指定月・全期間）ごとに別の名前で保存し、名前に「世代番号」を付ける。
// 書き込みのたびに世代番号を進めるので、古い世代のキャッシュは参照されなくなり60秒で自然に消える（3.2-1・3.5-7）。
const RECORDS_CACHE_KEY = 'getAll_cache_v2';
const RECORDS_CACHE_GEN_KEY = 'getAll_cache_gen_v2';
// 60秒キャッシュ。書き込み時は即破棄するため打刻の反映は遅れない。
// （スプレッドシートを直接手編集した場合のみ最大60秒反映が遅れる）
const RECORDS_CACHE_SECONDS = 60;
// 作り直し中の印（20秒で失効）と、他の呼び出しがキャッシュの完成を待つ最長時間（3.10-4）
const RECORDS_REBUILD_FLAG_SECONDS = 20;
const RECORDS_REBUILD_WAIT_MS = 3000;

function doGet(e) {
  const action = e.parameter.action;
  let ss;
  try { ss = getTimecardSs(); } catch (err) { return createJsonResponse({ status: 'error', message: String(err.message) }); }

  if (action === 'getAll') {
    // 範囲: months=2（直近2か月）/ month=YYYYMM（指定月）/ 指定なし（全期間・古い版向け）
    const range = resolveRecordRange(e.parameter);
    if (!range) return createJsonResponse({ status: 'error', message: 'Invalid range' });
    const gen = getCacheGeneration(true);
    let cached = getRecordsCache(range.key, gen);
    if (cached) {
      return ContentService.createTextOutput(cached)
        .setMimeType(ContentService.MimeType.JSON);
    }
    // ===== キャッシュ再構築の同時多発（スタンピード）防止（3.10-4 で作り替え） =====
    // Webアプリから呼ぶ getDocumentLock() は null になり、以前の「1つずつ並べる」処理は効いていなかったのでやめた。
    // 作り直し中の印（CacheService・20秒で失効）があれば、最大3秒だけキャッシュの完成を待ち、
    // 間に合わなければ自分で読む。打刻（doPost）はこの印を見ないので待たされない。
    const cache = CacheService.getScriptCache();
    const flagKey = 'getAll_rebuilding_' + range.key;
    let mine = false;
    if (cache.get(flagKey)) {
      const waitUntil = Date.now() + RECORDS_REBUILD_WAIT_MS;
      while (Date.now() < waitUntil) {
        Utilities.sleep(300);
        cached = getRecordsCache(range.key, getCacheGeneration(true));
        if (cached) {
          return ContentService.createTextOutput(cached)
            .setMimeType(ContentService.MimeType.JSON);
        }
      }
    } else {
      try { cache.put(flagKey, '1', RECORDS_REBUILD_FLAG_SECONDS); mine = true; } catch (err) {}
    }
    try {
      return getAllRecords(range);
    } finally {
      if (mine) { try { cache.remove(flagKey); } catch (err) {} }
    }
  } else if (action === 'getRecord') {
    // 名前を選んだ時点で、その人のその日の行だけを返す（3.1(4)）。キャッシュは使わずシートを直接読む。
    // 返事にも世代番号を付ける（端末は今までに見た最大より古ければ使わない。3.9-3）
    const one = getOneRecord(ss, e.parameter.staffName, e.parameter.date);
    one.generation = Number(getCacheGeneration(true)) || 0;
    return createJsonResponse(one);
  } else if (action === 'opStatus') {
    // POST の返事が読めない端末のための確認（3.6-4）: opId ごとの結果を返す
    // 1回に100件まで（3.8-8）
    const ids = String(e.parameter.ids || '').split(',').map(function(x) { return x.trim(); }).filter(Boolean).slice(0, 100);
    const log = loadOpLog(ss).map;
    const results = {};
    // 「受付」のまま（途中で止まった）は未完了なので返さない（端末は再送する）
    ids.forEach(function(id) {
      if (!log[id] || log[id].result === OPLOG_ACCEPTED) return;
      results[id] = { result: log[id].result, reason: log[id].reason };
      // 直した打刻の時刻（コードレビュー1-9）。最後に実際に書いた計画のもの（コードレビュー3-2）
      const cr = finalCorrection(log[id].before);
      if (cr) results[id].corrected = cr;
    });
    return createJsonResponse({ status: 'ok', results: results, generation: Number(getCacheGeneration(true)) || 0 });
  } else if (action === 'getSettings') {
    const settings = getSettings(ss);
    return createJsonResponse({ status: 'ok', settings: settings });
  } else if (action === 'ping' && e.parameter.clock) {
    // 時計のずれを測るための ping（3.11-3）: サーバー時刻を ms で返すだけ（シートは読まない）
    return createJsonResponse({ status: 'ok', serverTimeMs: Date.now() });
  } else if (action === 'reportDevice') {
    // 端末名と未送信の件数の知らせ（3.11-5）
    return createJsonResponse(reportDevice(e.parameter));
  } else if (action === 'removeDevice') {
    // 管理画面の「この端末を一覧から消す」（コードレビュー1-8）
    return createJsonResponse(removeDevice(e.parameter));
  } else if (action === 'deviceStatus') {
    return createJsonResponse({ status: 'ok', devices: listRegisteredDevices(), serverTimeMs: Date.now() });
  } else if (action === 'ping') {
    return createJsonResponse({ status: 'ok', message: 'pong', timestamp: new Date().toISOString() });
  } else if (action === 'getLeaveInfo') {
    // 【2026-11-08 オーナー決定】月末時点の計算はやめ、今までどおり管理簿の今の値（今日時点）を返す
    const year = e.parameter.year || String(new Date().getFullYear());
    const result = getLeaveInfo(year);
    return createJsonResponse({ status: 'ok', found: result.found, fileName: result.fileName, leaveInfo: result.leaveInfo, reason: result.reason || null });
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
const LEAVE_PERIOD_SHEET = '設定';     // 管理簿ファイルの「設定」シート
const LEAVE_PERIOD_START_CELL = 'C2'; // 期首
const LEAVE_PERIOD_END_CELL = 'C3';   // 期末
const LEAVE_NAME_COL = 2;      // B列: 名前
const LEAVE_NEXT_GRANT_COL = 9; // I列: 基準日（次回付与日）
const LEAVE_REMAINING_COL = 14; // N列: 有給残日数

// 見た目がそっくりな異体字（例: はしご高「髙」と標準の「高」）を吸収して名前を照合する。
// 管理簿側の正式表記（戸籍表記）はそのままで良く、ここでは比較用にのみ正規化する。
const KANJI_VARIANTS = [
  ['髙', '高'], // 髙(はしご高) -> 高
  ['﨑', '崎'], // 﨑(たつさき) -> 崎
  ['齋', '斎'], ['斉', '斎'], // 齋・斉 -> 斎
  ['邊', '辺'], ['邉', '辺'], // 邊・邉 -> 辺
  ['櫻', '桜'], // 櫻 -> 桜
];
function normalizeNameForMatch(name) {
  let s = String(name || '').trim();
  for (const [from, to] of KANJI_VARIANTS) s = s.split(from).join(to);
  return s;
}

function getLeaveInfo(year) {
  const result = {};
  try {
    // 管理簿ファイルの選び方（2026-11-08 コードレビュー6-2・6-3・7-1・7-2）:
    // 名前がちょうど「年次有給休暇管理簿＋4けたの年」のもの（「旧」「のコピー」「控え」などが付いたものは使わない）、
    // かつゴミ箱に入っていないもののうち、「設定」シートの期首(C2)〜期末(C3)に今日が入る1件だけを使う
    // （年度が5月〜翌4月でも、1〜4月に正しい年度を選べる）。
    // 期間に入るものが0件・2件以上、または読めない候補があるときは「取得できず」（found:false）にして、
    // 間違ったファイルの値は出さない。
    const folder = DriveApp.getFolderById(LEAVE_FOLDER_ID);
    const nameRe = new RegExp('^' + LEAVE_FILE_NAME_PREFIX + '\\d{4}$');
    const today = new Date();
    const byPeriod = [];
    const unreadable = [];
    const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
    while (files.hasNext()) {
      const f = files.next();
      const n = f.getName();
      if (!nameRe.test(n)) continue;
      if (f.isTrashed && f.isTrashed()) continue;
      try {
        const st = SpreadsheetApp.open(f).getSheetByName(LEAVE_PERIOD_SHEET);
        const ps = st ? st.getRange(LEAVE_PERIOD_START_CELL).getValue() : null;
        const pe = st ? st.getRange(LEAVE_PERIOD_END_CELL).getValue() : null;
        if (!(ps instanceof Date) || !(pe instanceof Date)) { unreadable.push(n + '（期首・期末が読めない）'); continue; }
        const endOfDay = new Date(pe.getFullYear(), pe.getMonth(), pe.getDate(), 23, 59, 59).getTime();
        if (ps.getTime() <= today.getTime() && today.getTime() <= endOfDay) byPeriod.push(f);
      } catch (err) { unreadable.push(n + '（開けない）'); }
    }
    let file = null;
    let why = '';
    if (unreadable.length > 0) why = '読めない管理簿があります: ' + unreadable.join(', ');
    else if (byPeriod.length > 1) why = '今日が期間に入る管理簿が複数あります: ' + byPeriod.map(function(f) { return f.getName(); }).join(', ');
    else if (byPeriod.length === 0) why = '今日が期間に入る管理簿が見つかりません';
    else file = byPeriod[0];
    if (!file) {
      // アプリ側でアラート表示できるよう found:false と理由を返す（打刻機能自体は止めない）
      console.error('getLeaveInfo: ' + why);
      return { found: false, fileName: null, leaveInfo: result, reason: why };
    }
    const leaveSs = SpreadsheetApp.open(file);
    const sheet = leaveSs.getSheetByName(LEAVE_SHEET_NAME);
    if (!sheet) return { found: false, fileName: file.getName(), leaveInfo: result };
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { found: true, fileName: file.getName(), leaveInfo: result };
    const values = sheet.getRange(2, 1, lastRow - 1, LEAVE_REMAINING_COL).getValues();
    for (const row of values) {
      const rawName = String(row[LEAVE_NAME_COL - 1] || '').trim();
      if (!rawName) continue;
      const name = normalizeNameForMatch(rawName);
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
  let ss;
  try { ss = getTimecardSs(); } catch (err) { return createJsonResponse({ status: 'error', message: String(err.message) }); }
  let postData;
  try {
    postData = JSON.parse(e.postData.contents);
  } catch (err) {
    return createJsonResponse({ status: 'error', message: 'Invalid JSON' });
  }

  const action = postData.action;

  // ===== 版の印（3.5-5） =====
  // 古いアプリの勤怠の書き込み（record / sync / delete / confirmMonth）は断る。
  // 名簿・PIN の保存（saveSettings）は今のまま（独自の安全策がある）。
  if (action === 'record' || action === 'sync' || action === 'delete' || action === 'confirmMonth') {
    return createJsonResponse({ status: 'error', message: 'old_client' });
  }
  if (action === 'ops' && !(Number(postData.clientVersion) >= MIN_CLIENT_VERSION)) {
    return createJsonResponse({ status: 'error', message: 'old_client' });
  }

  if (action === 'ops') return handleOps(ss, postData);

  // ===== 排他制御（ロック機能）の追加 =====
  // 複数のPCからの通信が重なった際、最大10秒間待機して順番に確実に処理する
  // （順番待ちを超えた場合も HTTP は 200 で status:'error' を返す。アプリは中身で判定して再送する）
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (err) {
    return createJsonResponse({ status: 'error', message: 'Server busy. Please try again.' });
  }

  try {
    if (action === 'saveSettings') {
      const result = saveSettings(ss, postData.settings);
      return createJsonResponse(result);
    }
  } finally {
    lock.releaseLock();
  }

  return createJsonResponse({ status: 'error', message: 'Invalid action' });
}

// ===== 操作（ops）の受付 =====
// 【コードレビュー1-11】スクリプトのロックは「操作1つごと」に取って放す（まとめて届いた未送信の途中でも、
// もう1台の打刻が割り込める）。書く順番（受付→flush→勤怠・EditLog→flush→完了→flush）は操作ごとに今までどおり。
// ロックを取るたびに OpLog の増えた行と、この opId の行を読み直す（他の実行が同じ opId を先に処理していても二重に当てない）。
function handleOps(ss, postData) {
  const ops = Array.isArray(postData.ops) ? postData.ops : [];
  const sender = String(postData.sender || '').substring(0, 40);
  const receivedAt = Date.now();
  // 時計のずれの目安（警告・EditLog 用）: 受信時刻－端末の送信時刻(sentAt)。通信の遅れも含むので、
  // 操作の時刻を直すのには使わない（直すのは操作に付いた offsetAtOp。3.11-3）
  const sentAt = Number(postData.sentAt || postData.clientNow);
  const skewMs = isFinite(sentAt) && sentAt > 0 ? receivedAt - sentAt : 0;
  const opLog = { map: {}, lastRows: {}, loaded: false };
  const results = {};
  let processed = 0;
  let busy = false;
  let skewLogged = Math.abs(skewMs) < CLOCK_SKEW_WARN_MS;
  ops.forEach(function(op) {
    const opId = op && typeof op.opId === 'string' ? op.opId : '';
    if (!opId) { results[''] = { result: 'invalid', reason: 'opIdがありません' }; return; }
    // 順番待ちを超えたら、残りの操作は待たずに「一時エラー」で返す（端末が再送する）
    if (busy) { results[opId] = { result: 'error', reason: 'busy' }; return; }
    const lock = LockService.getScriptLock();
    try { lock.waitLock(10000); } catch (err) { busy = true; results[opId] = { result: 'error', reason: 'busy' }; return; }
    try {
      if (!skewLogged) { skewLogged = true; logClockSkew(ss, sender, skewMs); }
      refreshOpLog(ss, opLog, opId);
      results[opId] = processOneOp(ss, op, opId, sender, Date.now(), opLog);
      processed++;
    } finally {
      lock.releaseLock();
    }
  });
  if (processed === 0 && busy) return createJsonResponse({ status: 'error', message: 'Server busy. Please try again.' });
  // 書き込みの返事にも世代番号を入れる（3.8-5）
  const out = { status: 'ok', results: results, serverTime: new Date(receivedAt).toISOString(), generation: Number(getCacheGeneration(true)) || 0 };
  out.skewMs = skewMs;
  if (Math.abs(skewMs) >= CLOCK_SKEW_WARN_MS) out.clockSkewMs = skewMs;
  return createJsonResponse(out);
}

// 操作1つを当てる（ロックの中で呼ぶ）
function processOneOp(ss, op, opId, sender, serverNow, opLog) {
  const logged = opLog.map[opId];
  // 同じ opId がもう一度届いたら、書き込まずに最初の結果を返す（3.6-4）
  if (logged && logged.result !== OPLOG_ACCEPTED) {
    const dup = { result: logged.result, reason: logged.reason, duplicate: true };
    const cr = finalCorrection(logged.before);
    if (cr) dup.corrected = cr;
    return dup;
  }
  let normAt, seq, logRow = null;
  // 当て直し（3.12-1）: 「受付」に残した変更前の値・書く値と「EditLog 済み」の印を使う
  const ctx = { resumed: !!logged, editLogged: !!(logged && logged.editLogged), stored: logged ? logged.before : null, opLog: opLog };
  if (logged) {
    // 「受付」のまま残っていた（前回が途中で止まった）: 記録してある比較時刻・受付順でもう一度当てる（3.9-2）
    normAt = logged.normAt; seq = logged.seq; logRow = logged.ref;
    // 前回の書き込みが届いていなかったとき（コードレビュー2-2）: 最初に残した変更前の値・書く値は上書きしない。
    // 今回あらためて決めた値は、同じ欄の中の re に足して残す（2回目に止まっても、最後の計画から EditLog を作れる）
    ctx.accept = function(before) {
      const first = logged.before || {};
      const prevRe = first.re ? JSON.stringify(first.re) : '';
      const sameAsFirst = JSON.stringify(Object.assign({}, first, { re: undefined })) === JSON.stringify(before);
      if (logRow && logged.before && !sameAsFirst && prevRe !== JSON.stringify(before)) {
        const merged = Object.assign({}, first, { re: before });
        getOpLogSheet(ss, logRow.sheet, true).getRange(logRow.row, 12).setValue(JSON.stringify(merged));
        SpreadsheetApp.flush();
        logged.before = merged;
      }
      return before;
    };
  } else {
    // 比べる時刻と受付順は、最初に受け付けたときに決めて OpLog に残す（3.8-2・3.9-2）
    normAt = normalizeOpTime(op, serverNow);
    seq = nextOpSeq();
    if (normAt < serverNow - OP_TOO_OLD_DAYS * 86400000) {
      // OpLog に無く、180日より古い操作は当てない（3.8-8）
      const r0 = { result: 'skipped', reason: 'too_old' };
      appendOpLogRow(ss, opLogRowValues(op, r0.result, r0.reason, serverNow, sender, normAt, seq), serverNow);
      opLog.map[opId] = { result: r0.result, reason: r0.reason };
      return r0;
    }
    // ① 書く内容を決めたら（読むだけ）、勤怠を書く前に「受付」を書く。変更前の値・書く値も一緒に残す（3.9-2・3.12-1）
    ctx.accept = function(before) {
      logRow = appendOpLogRow(ss, opLogRowValues(op, OPLOG_ACCEPTED, '', serverNow, sender, normAt, seq, before), serverNow);
      opLog.map[opId] = { result: OPLOG_ACCEPTED, reason: '', normAt: normAt, seq: seq, ref: logRow, before: before, staff: String(op.staffName || ''), date: String(op.date || '') };
      return before;
    };
  }
  // ② 勤怠の行と EditLog を書く（1日分の1〜17列を1回で）。EditLog を書けたら「EditLog 済み」の印を付ける
  ctx.markEditLogged = function() { if (logRow) markOpLogEditLogged(ss, logRow); };
  let res;
  try {
    res = applyOp(ss, op, sender, serverNow, normAt, seq, ctx);
  } catch (err) {
    console.error('applyOp failed', err);
    res = { result: 'error', reason: String(err && err.message ? err.message : err) };
  }
  // 勤怠の行を書いたらすぐ世代を進める（EditLog の失敗で「受付」のまま残るときも。書いた分のキャッシュは使わない）
  if (res.wrote || res.result === 'applied') bumpCacheGeneration();
  delete res.wrote;
  // ③ OpLog を「完了」にする。error（一時的な失敗）は「受付」のまま残す＝再送で同じ時刻・受付順で当て直す
  if (res.result !== 'error') {
    if (logRow) finishOpLogRow(ss, logRow, res.result, res.reason);
    else appendOpLogRow(ss, opLogRowValues(op, res.result, res.reason, serverNow, sender, normAt, seq), serverNow); // 書く前に決まった結果（invalid など）
    opLog.map[opId] = { result: res.result, reason: res.reason || '', before: (opLog.map[opId] || {}).before, seq: seq, normAt: normAt, staff: String(op.staffName || ''), date: String(op.date || ''), ref: logRow };
  }
  return res;
}

// ===== 端末の未送信の知らせ（3.11-5・3.12-4） =====
// 各端末は同期のたびに（変わったとき・数分おき）端末名と未送信の件数を知らせる。
// 管理画面で名前を付けた端末を「登録端末」とし、deviceStatus は登録端末すべてを「最後の知らせからの経過」付きで返す。
// （10分以内・未送信0件かどうかはアプリが判断する）
const DEVICE_PROP_PREFIX = 'DEV_';
function reportDevice(p) {
  const id = String(p.id || '').replace(/[^\w\-]/g, '').substring(0, 40);
  if (!id) return { status: 'error', message: 'id がありません' };
  const pending = Math.max(0, parseInt(p.pending, 10) || 0);
  const info = { name: String(p.name || '').substring(0, 40), pending: pending, at: Date.now() };
  PropertiesService.getScriptProperties().setProperty(DEVICE_PROP_PREFIX + id, JSON.stringify(info));
  return { status: 'ok' };
}
// 【コードレビュー1-8】知らせのあった端末は、名前が無くてもすべて「知っている端末」として返す
// （名前を付け忘れた端末も、エクセル出力前の確認から漏れない）。
// 【コードレビュー2-3】長く知らせの無い端末も外さない（3.12-4「登録端末すべて」）。管理者が管理画面の「一覧から消す」
// （removeDevice）で消すまで、出力前の確認に残る。最後の知らせの日時（at・ageMs）は今までどおり返す。
function listRegisteredDevices() {
  const all = PropertiesService.getScriptProperties().getProperties();
  const now = Date.now();
  const out = [];
  Object.keys(all).forEach(function(k) {
    if (k.indexOf(DEVICE_PROP_PREFIX) !== 0) return;
    let v = null;
    try { v = JSON.parse(all[k]); } catch (e) { return; }
    if (!v || !isFinite(Number(v.at))) return;
    const ageMs = now - Number(v.at);
    out.push({ id: k.substring(DEVICE_PROP_PREFIX.length), name: String(v.name || ''), pending: Number(v.pending) || 0, at: new Date(Number(v.at)).toISOString(), ageMs: ageMs });
  });
  return out;
}
function removeDevice(p) {
  const id = String(p.id || '').replace(/[^\w\-]/g, '').substring(0, 40);
  if (!id) return { status: 'error', message: 'id がありません' };
  PropertiesService.getScriptProperties().deleteProperty(DEVICE_PROP_PREFIX + id);
  return { status: 'ok' };
}

// 比べる時刻（3.11-3・3.12-2）: 端末が ping で測ったずれ（操作を作ったときの offsetAtOp＝サーバー時刻－端末時刻）があれば
// 常に at＋offsetAtOp。測れていない操作（起動直後・オフライン）は at のまま。
// （2分を超えるときだけ直すのは「記録する打刻の時刻」だけ: applyOp を参照）
function normalizeOpTime(op, serverNow) {
  const atParsed = op ? Date.parse(op.at) : NaN;
  if (!isFinite(atParsed)) return serverNow;
  const off = op && op.offsetAtOp !== undefined && op.offsetAtOp !== null && op.offsetAtOp !== '' ? Number(op.offsetAtOp) : NaN;
  return isFinite(off) ? atParsed + off : atParsed;
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===== 読み込み範囲 =====
// months=2 → 今月と先月（スクリプトのタイムゾーン）/ month=YYYYMM → その月だけ / 指定なし → 全期間
function resolveRecordRange(params) {
  params = params || {};
  if (params.month) {
    const m = String(params.month).replace('-', '');
    if (!/^\d{6}$/.test(m)) return null;
    return { key: 'm' + m, months: [m] };
  }
  if (params.months) {
    const n = Math.max(1, Math.min(12, parseInt(params.months, 10) || 2));
    const tz = Session.getScriptTimeZone();
    const now = new Date();
    const y = parseInt(Utilities.formatDate(now, tz, 'yyyy'), 10);
    const mo = parseInt(Utilities.formatDate(now, tz, 'MM'), 10);
    const list = [];
    for (let i = 0; i < n; i++) {
      const d = new Date(y, mo - 1 - i, 1);
      list.push(String(d.getFullYear()) + String(d.getMonth() + 1).padStart(2, '0'));
    }
    return { key: 'n' + n + '_' + list[0], months: list };
  }
  return { key: 'all', months: null };
}

// ===== 全レコード読み取り =====
function getAllRecords(range) {
  const ss = getTimecardSs();
  const sheets = ss.getSheets();
  const tz = Session.getScriptTimeZone();
  range = range || { key: 'all', months: null };
  // 読み込み開始時の世代番号を覚えておき、終わったときに変わっていたらキャッシュに保存しない（3.2-1）
  const genAtStart = getCacheGeneration(true);

  // IDをキーとしたMapを使い、同一IDで複数行ある場合は最新を残す
  // clientUpdatedAt (11列目) または serverUpdatedAtMs (10列目) を優先
  const recordMap = {};
  const unreadableSheets = [];

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    // シート名は「スタッフ名_YYYYMM」形式を期待
    const nm = sheetName.match(/^(.+)_(\d{6})$/);
    if (!nm || nm[1] === OPLOG_SHEET_NAME) return; // OpLog_YYYYMM は勤怠のシートではない
    if (range.months && range.months.indexOf(nm[2]) === -1) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    // 17列まで取得（日付、曜日、出勤、中抜け、退勤、賄い、有給、備考、更新日時、clientUpdatedAt、削除フラグ、休憩分数、交通費、送信元、予備、項目メタ）
    // ただし、最大列数(getMaxColumns)を超えて読み込もうとすると境界外エラーになるため、安全な範囲を指定する
    let data;
    try {
      data = readSheetRows(sheet);
    } catch (e) {
      // 【2026-11-08】読めなかったシートは黙って飛ばさず、返事に入れる（エクセル出力はこれを見て止める）
      console.warn('Sheet access error: ' + sheetName, e);
      unreadableSheets.push(sheetName);
      return;
    }

    const staffName = nm[1];
    data.forEach(row => {
      const dateStr = cellToDateStr(row[0], tz);
      if (!dateStr) return;
      const record = rowToRecord(row, staffName, dateStr, tz);
      const compareMs = rowCompareMs(row);
      // 同じIDが複数行ある場合（重複行）は最新のものだけ残す
      if (!recordMap[record.id] || compareMs > recordMap[record.id]._compareMs) {
        record._compareMs = compareMs;
        recordMap[record.id] = record;
      }
    });
  });

  // 比較用フィールドを除いて返す
  const allRecords = Object.values(recordMap).map(r => {
    const rec = Object.assign({}, r);
    delete rec._compareMs;
    return rec;
  });

  const genAtEnd = getCacheGeneration(false);
  const stale = genAtEnd !== genAtStart;
  const payload = JSON.stringify({
    status: 'ok',
    count: allRecords.length,
    records: allRecords,
    range: range.months || 'all',
    // 端末は、すでに反映したものより古い返事を捨てる（3.7-6）
    generation: Number(genAtStart) || 0,
    unreadableSheets: unreadableSheets,
    stale: stale,
    timestamp: new Date().toISOString()
  });

  // 結果を短時間キャッシュ（保存できない場合は黙ってスキップ＝従来動作）
  // 読めないシートがあった結果・途中で書き込みがあった結果は保存しない
  if (!stale && unreadableSheets.length === 0) putRecordsCache(range.key, genAtStart, payload);

  return ContentService.createTextOutput(payload)
    .setMimeType(ContentService.MimeType.JSON);
}

// シートの2行目以降を17列まで読む（無い列は空扱い）
function readSheetRows(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const maxCol = sheet.getMaxColumns();
  const colsToRead = Math.min(Math.max(sheet.getLastColumn(), 1), maxCol, SHEET_COLS);
  const values = sheet.getRange(2, 1, lastRow - 1, colsToRead).getValues();
  return values.map(function(r) {
    const row = r.slice();
    while (row.length < SHEET_COLS) row.push('');
    return row;
  });
}

function cellToDateStr(dateCell, tz) {
  if (!dateCell) return '';
  try {
    if (dateCell instanceof Date) return Utilities.formatDate(dateCell, tz, 'yyyy-MM-dd');
    if (typeof dateCell === 'string') {
      const match = dateCell.match(/(\d{4}-\d{2}-\d{2})/);
      return match ? match[1] : '';
    }
  } catch (e) { /* 日付パース失敗時はスキップ */ }
  return '';
}

// 重複比較用: clientUpdatedAt > サーバー側updatedAt の優先順
function rowCompareMs(row) {
  try {
    const cu = row[COL.clientUpdatedAt - 1] ? String(row[COL.clientUpdatedAt - 1]) : '';
    if (cu && !isNaN(new Date(cu).getTime())) return new Date(cu).getTime();
    if (row[COL.serverUpdatedAt - 1] instanceof Date) return row[COL.serverUpdatedAt - 1].getTime();
  } catch (e) {}
  return 0;
}

function parseMeta(cell) {
  let meta = {};
  if (cell) {
    try { meta = JSON.parse(String(cell)) || {}; } catch (e) { meta = {}; }
  }
  // f: 項目ごとの記録 { 項目: { t: 最後に書いた操作の時刻(ms), m: 手修正済みなら1 } }（3.6-2）
  // 17列目が空の既存行は「記録なし（時刻0・印なし）」として扱う
  if (!meta.f || typeof meta.f !== 'object') meta.f = {};
  return meta;
}

// 項目ごとの記録を読む（3.7-3）。17列目が空の既存行は、時刻の初期値＝11列目（無ければ10列目）、
// 備考に【本人修正済】がある既存行は時刻欄を手修正済みとみなす。
// 返り値の f は { t, m(手修正済み), L(有給による空欄) }
function fieldInfo(meta, f) {
  const x = meta.f[f];
  return { t: (x && Number(x.t)) || 0, q: (x && Number(x.q)) || 0, m: !!(x && x.m), L: !!(x && x.L) };
}
const ALL_FIELDS = ['clockIn', 'breakStart', 'breakEnd', 'clockOut', 'meal', 'transport', 'remarks', 'additionalBreakMins', 'isPaidLeave', 'deleted'];
// 17列目が「使える」か: 空でなく、かつ 10列目（サーバーの更新時刻）と17列目に残した保存時刻（savedAt）がぴったり同じこと（3.11-1・3.12-3）。
// 新しい版が保存した行は10列目と savedAt が同じなので、遅れて届いた打刻を保存した行も誤判定しない。
// 【3.10-6】古い版の GAS に戻していた間に書かれた行は15〜17列目が更新されず値と食い違うので、17列目を空として扱い、
// 最初に書くときに作り直す。
function hasMetaCell(row) {
  const raw = String(row[COL.meta - 1] || '').trim();
  if (!raw) return false;
  const meta = parseMeta(raw);
  const su = row[COL.serverUpdatedAt - 1];
  const suMs = su instanceof Date ? su.getTime() : (su ? Date.parse(String(su)) : NaN);
  // 新しい版は10列目と savedAt に同じ値（秒にそろえた時刻）を書く。秒の単位で一致しなければ古い版で書かれた行（3.12-3）。
  // 【コードレビュー1-6】スプレッドシートの日時は小数で持つため、読み出すと数ミリ秒ずれることがある → 秒にそろえて比べる
  if (!meta.savedAt || !isFinite(suMs) || Math.round(suMs / 1000) !== Math.round(Number(meta.savedAt) / 1000)) return false;
  return true;
}
function metaForRow(row) {
  const valid = hasMetaCell(row);
  const meta = valid ? parseMeta(row[COL.meta - 1]) : { f: {} };
  if (!valid) {
    const base = rowCompareMs(row);
    const tagged = String(row[COL.remarks - 1] || '').indexOf('【本人修正済】') !== -1;
    ALL_FIELDS.forEach(function(f) {
      const e = { t: base };
      if (tagged && TIME_FIELDS.indexOf(f) !== -1) e.m = 1;
      meta.f[f] = e;
    });
  }
  return meta;
}
// 有給かどうか（3.7-4）: 17列目のある行は8列目だけ。17列目の無い既存行だけ備考の「有給申請」も見る
function rowIsPaidLeave(row) {
  if (row[COL.isPaidLeave - 1] === '有給') return true;
  if (hasMetaCell(row)) return false;
  const remarks = row[COL.remarks - 1] ? String(row[COL.remarks - 1]) : '';
  return remarks.indexOf('有給申請') !== -1;
}

function rowToRecord(row, staffName, dateStr, tz) {
  const clientUpdatedAt = row[COL.clientUpdatedAt - 1] ? String(row[COL.clientUpdatedAt - 1]) : '';
  const remarks = row[COL.remarks - 1] ? String(row[COL.remarks - 1]) : '';
  // 列13（休憩分数）: 空セルはnullを返す（旧データと明示的な0を区別するため）
  let additionalBreakMins = null;
  const bm = row[COL.additionalBreakMins - 1];
  if (bm !== '' && bm !== null && bm !== undefined) {
    const parsed = parseInt(bm);
    additionalBreakMins = isNaN(parsed) ? null : parsed;
  }
  // 列14（交通費）: 空セルはnull
  let transport = null;
  const tp = row[COL.transport - 1];
  if (tp !== '' && tp !== null && tp !== undefined) {
    const t = parseInt(tp);
    transport = isNaN(t) ? null : t;
  }
  const isActuallyPaidLeave = rowIsPaidLeave(row);
  const mealCell = row[COL.meal - 1];
  const delCell = row[COL.deleted - 1];
  const meta = metaForRow(row);
  return {
    id: staffName + '_' + dateStr,
    staffName: staffName,
    date: dateStr,
    clockIn: formatTimeCell(row[COL.clockIn - 1]) || null,
    clockOut: formatTimeCell(row[COL.clockOut - 1]) || null,
    breakStart: formatTimeCell(row[COL.breakStart - 1]) || null,
    breakEnd: formatTimeCell(row[COL.breakEnd - 1]) || null,
    meal: mealCell === '有' || mealCell === 1 || mealCell === true,
    isPaidLeave: isActuallyPaidLeave,
    remarks: remarks,
    additionalBreakMins: additionalBreakMins,
    transport: transport,
    clientUpdatedAt: clientUpdatedAt || null,
    serverUpdatedAtMs: (row[COL.serverUpdatedAt - 1] instanceof Date) ? row[COL.serverUpdatedAt - 1].getTime() : 0,
    // 12列目（index 11）= deletedフラグ
    deleted: delCell === 'true' || delCell === true,
    // 【2026-11-08】16列目: 本人修正の印（3.6-3）。17列目: 項目ごとの操作時刻・手修正済み印（3.6-2・3.6-5 の合流に使う）
    selfEdited: row[COL.selfEdited - 1] === '本人修正',
    fields: meta.f
  };
}

// ===== 1件取得（getRecord） =====
function getOneRecord(ss, staffName, dateStr) {
  if (!staffName || !/^\d{4}-\d{2}-\d{2}$/.test(String(dateStr || ''))) {
    return { status: 'error', message: 'スタッフ名と日付を指定してください' };
  }
  const tz = Session.getScriptTimeZone();
  const sheet = ss.getSheetByName(staffName + '_' + dateStr.substring(0, 4) + dateStr.substring(5, 7));
  if (!sheet) return { status: 'ok', record: null };
  const found = findRowForDate(sheet, dateStr, tz);
  if (found.rowIndex === -1) return { status: 'ok', record: null };
  return { status: 'ok', record: rowToRecord(found.row, staffName, dateStr, tz) };
}

// ===== キャッシュ分割保存 =====
// CacheServiceは1キーあたり100KBまでしか保存できないため、
// 応答JSON（現在135KB超）を分割して保存・復元する（1つ 90KB 以下）。
function cacheBaseKey(rangeKey, gen) {
  return RECORDS_CACHE_KEY + '_' + rangeKey + '_' + gen;
}

function putRecordsCache(rangeKey, gen, payload) {
  if (!gen) return;
  try {
    const cache = CacheService.getScriptCache();
    const base = cacheBaseKey(rangeKey, gen);
    // 【コードレビュー1-7】100KB の上限は UTF-8 のバイト数。1文字（UTF-16 の1単位）は最大3バイトなので、
    // 30,000 文字ずつに分ければ 1つ 90KB 以下になる（日本語ばかりでも超えない）。サロゲートペアは途中で切らない。
    const CHUNK_SIZE = 30000;
    const entries = {};
    let n = 0;
    let i = 0;
    while (i < payload.length) {
      let end = Math.min(i + CHUNK_SIZE, payload.length);
      const last = payload.charCodeAt(end - 1);
      if (end < payload.length && last >= 0xD800 && last <= 0xDBFF) end--;
      entries[base + '_' + n] = payload.substring(i, end);
      i = end;
      n++;
    }
    entries[base + '_meta'] = String(n);
    cache.putAll(entries, RECORDS_CACHE_SECONDS);
  } catch (err) {
    // キャッシュ保存失敗は無視（動作に影響なし）
  }
}

function getRecordsCache(rangeKey, gen) {
  if (!gen) return null;
  try {
    const cache = CacheService.getScriptCache();
    const base = cacheBaseKey(rangeKey, gen);
    const meta = cache.get(base + '_meta');
    if (!meta) return null;
    const n = parseInt(meta);
    if (!n || n < 1) return null;
    const keys = [];
    for (let i = 0; i < n; i++) keys.push(base + '_' + i);
    const parts = cache.getAll(keys);
    let out = '';
    for (let i = 0; i < n; i++) {
      const p = parts[base + '_' + i];
      if (p === undefined || p === null) return null; // チャンク欠損時はキャッシュ不成立として通常処理へ
      out += p;
    }
    return out;
  } catch (err) {
    return null;
  }
}

// 世代番号（3.5-7）: CacheService に置く。無くなっていたら新しい番号を作る（＝キャッシュを使わないだけで安全側）
// 世代番号はミリ秒単位の現在時刻（3.7-8）。CacheService から消えても、新しく作る番号が古い番号と重ならない
function newCacheGeneration(prev) {
  return String(Math.max(Date.now(), (Number(prev) || 0) + 1));
}
function getCacheGeneration(createIfMissing) {
  try {
    const cache = CacheService.getScriptCache();
    let gen = cache.get(RECORDS_CACHE_GEN_KEY);
    if (!gen && createIfMissing) {
      gen = newCacheGeneration();
      cache.put(RECORDS_CACHE_GEN_KEY, gen, 21600);
    }
    return gen || '';
  } catch (err) {
    return '';
  }
}

// 書き込み発生時に世代番号を進める（古い世代のキャッシュは参照されなくなる）
function bumpCacheGeneration() {
  try {
    const cache = CacheService.getScriptCache();
    cache.put(RECORDS_CACHE_GEN_KEY, newCacheGeneration(cache.get(RECORDS_CACHE_GEN_KEY)), 21600);
  } catch (err) {
    // 進められなくても古いキャッシュは最大60秒で自然消滅するため致命的ではない
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

// ===== 操作（op）の書き込み（設計書 3.1(2)・3.5・3.6・3.7） =====
// op = { opId, type, staffName, date, at(端末で操作した時刻), data }
// 返り値 = { result: 'applied' | 'skipped' | 'invalid' | 'error', reason? }
// どの項目を書くかの判断（ルール）は下の planXxx 関数にまとめてある。ルール変更はそこだけ直せばよい。
// 新旧は「操作した時刻をサーバーの時計に直した時刻（at＋差）」で決める（3.6-1・3.7-1）。
// 直した時刻が全く同じなら、後から受け付けた方を新しいとする（比較は「以上」）。

function applyOp(ss, op, sender, serverNow, normAt, seq, ctx) {
  ctx = ctx || { accept: function(b) { return b; }, markEditLogged: function() {}, editLogged: false };
  if (!op || typeof op !== 'object' || !op.opId || typeof op.opId !== 'string') {
    return { result: 'invalid', reason: 'opIdがありません' };
  }
  const atParsed0 = Date.parse(op.at);
  // 比べる時刻は OpLog に記録した「直した操作時刻」（3.8-2・3.9-2）
  const atConfirm = isFinite(Number(normAt)) ? Number(normAt) : (isFinite(atParsed0) ? atParsed0 : (serverNow || Date.now()));
  if (op.type === 'confirm') {
    const d = op.data || {};
    // 確認日時は実行時刻ではなく操作時刻（直したもの）を使う（3.8-3）
    ctx.accept({ exists: false, v: {} });
    const r = confirmMonth(ss, op.staffName, d.yearMonth, op.opId, atConfirm, seq);
    return { result: r.result || (r.status === 'ok' ? 'applied' : 'invalid'), reason: r.message || r.reason };
  }
  const planner = OP_PLANNERS[op.type];
  if (!planner) return { result: 'invalid', reason: '不明な操作: ' + op.type };
  const staffName = String(op.staffName || '');
  let dateStr = String(op.date || '');
  if (!staffName || !/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
    return { result: 'invalid', reason: 'スタッフ名または日付が不正です' };
  }
  const at = atConfirm;
  let opData = op.data || {};
  // 打刻の時刻そのもの（3.8-2）: 差が2分以内なら端末が表示した時刻のまま。超えるときは直した時刻（と日付）で記録し、EditLog に残す
  let clockFix = null;
  if (op.type === 'punch' && opData.time && isFinite(atParsed0) && Math.abs(at - atParsed0) > CLOCK_SKEW_WARN_MS) {
    const tzFix = Session.getScriptTimeZone();
    const fixed = new Date(at);
    const newDate = Utilities.formatDate(fixed, tzFix, 'yyyy-MM-dd');
    const newTime = Utilities.formatDate(fixed, tzFix, 'HH:mm:ss');
    clockFix = { from: dateStr + ' ' + normalizeTimeValue(opData.time), to: newDate + ' ' + newTime };
    dateStr = newDate;
    opData = Object.assign({}, opData, { time: newTime });
  }

  const tz = Session.getScriptTimeZone();
  const sheetName = staffName + '_' + dateStr.substring(0, 4) + dateStr.substring(5, 7);
  let sheet = ss.getSheetByName(sheetName);
  const found = sheet ? findRowForDate(sheet, dateStr, tz) : { rowIndex: -1, row: null, lastRow: 1 };
  const exists = found.rowIndex !== -1;
  const oldRow = exists ? found.row.slice() : blankRow(dateStr);
  const meta = metaForRow(oldRow);

  // 打刻の時刻を直したときは、端末の画面も直せるように返事に入れる（コードレビュー1-9）
  const corrected = clockFix ? { date: dateStr, field: opData.field, time: opData.time, fromDate: String(op.date || ''), written: true } : undefined;

  // 【コードレビュー1-1】当て直し: 前回の勤怠の書き込みが届いていたか（17列目の項目にこの操作の時刻と受付順があるか）。
  // 1回で1〜17列を書くので、1項目でも残っていれば書けている。届いていれば書き直さず、EditLog を（まだなら）
  // 「受付」に残した変更前の値と書く値から作って、最初に決めた結果（applied）を返す。
  // 【コードレビュー2-2】その項目を後から別の操作が書き換えていても、「その後で最初にその項目を書いた操作」が
  // 受け付けたときの変更前の値が、この操作の書く値と同じなら、この操作の書き込みは届いていた（その操作はこの値を見て書いた）。
  // 当て直しで最後に決めた値（re）があればそれを、無ければ最初に残した値を使う。
  const storedAll = ctx.resumed && ctx.stored && ctx.stored.w ? ctx.stored : null;
  const stored = storedAll ? effectivePlan(storedAll) : null;
  if (stored && Number(seq) > 0) {
    const ownMark = Object.keys(stored.w).some(function(f) {
      const x = meta.f[f];
      return !!x && Number(x.t) === at && Number(x.q) === Number(seq);
    });
    const landed = ownMark || landedByWriteChain(ss, ctx.opLog, Number(seq), stored, meta);
    if (landed) {
      if (!ctx.editLogged) {
        try {
          writeEditLogForOp(ss, sheetName, staffName, dateStr, oldRow, stored, stored.w, Object.keys(stored.w), stored.cf || null, stored.cff || null, op.opId, sender, true);
          ctx.markEditLogged();
        } catch (e) {
          console.warn('EditLog failed:', e);
          return { result: 'error', reason: 'editlog_failed' };
        }
        SpreadsheetApp.flush();
      }
      return { result: 'applied', reason: stored.r, corrected: finalCorrection(stored) };
    }
  }

  const state = rowState(oldRow, meta, seq);
  // 操作時刻が削除より前の操作は以後すべて skipped（3.6-2）
  const del = fieldInfo(meta, 'deleted');
  let plan;
  if (state.deleted && del.t && !state.notOlder(at, del)) plan = { writes: {}, reason: 'before_delete' };
  else plan = planner(state, at, opData);
  const writes = plan.writes || {};
  const fields = Object.keys(writes);
  // 変更前の値（3.12-1）: 書き換える項目の今の値。OpLog「受付」に残し、EditLog はこれと書いた値から作る。
  // 当て直しのときは、最初に受け付けたときの変更前の値が返る
  // 【コードレビュー1-1】書く値（w）・理由（r）・時計補正（cf / c）も一緒に残し、当て直しで EditLog を作れるようにする
  const beforeNow = { exists: exists, v: {} };
  fields.forEach(function(f) {
    const v = oldRow[COL[f] - 1];
    beforeNow.v[f] = TIME_FIELDS.indexOf(f) !== -1 ? (formatTimeCell(v) || '') : (v instanceof Date ? v.toISOString() : v);
  });
  // 【コードレビュー2-4】直した時刻を返すのは、この操作が時刻の欄を本当に書くときだけ（休憩分数だけ書くときは返さない）
  const timeWritten = !!(clockFix && opData.field && Object.prototype.hasOwnProperty.call(writes, opData.field));
  const correctedOut = timeWritten ? corrected : undefined;
  if (fields.length > 0) {
    beforeNow.w = writes;
    beforeNow.d = dateStr; // 実際に書いた日付（時計補正で日付が変わることがある）
    // 【コードレビュー3-1】書く項目ごとに「今その値を書いた操作の受付順」（直前に書いた操作）を残す。
    // 項目の書き込みの順番そのものの鎖になり、受付順と書いた順が逆でも「誰が誰の上に書いたか」をたどれる
    beforeNow.pq = {};
    fields.forEach(function(f) { beforeNow.pq[f] = (meta.f[f] && Number(meta.f[f].q)) || 0; });
    if (plan.reason) beforeNow.r = plan.reason;
    if (clockFix) { beforeNow.cf = clockFix; beforeNow.cff = opData.field; }
    if (correctedOut) beforeNow.c = correctedOut;
    else if (clockFix) beforeNow.c = { written: false, field: opData.field }; // 時刻の欄は書かない（休憩分数だけなど。コードレビュー3-2）
  }
  const before = ctx.accept(beforeNow) || beforeNow;
  const result = fields.length > 0 ? 'applied' : 'skipped';
  if (result === 'skipped') return { result: 'skipped', reason: plan.reason || 'nothing_to_write' };

  if (!sheet) sheet = createStaffSheet(ss, sheetName);
  else ensureSheetColumns(sheet);

  // 10列目と savedAt に同じ値を書く（秒にそろえる: スプレッドシートの日時でも値が変わらないように。3.12-3）
  const now = new Date(Math.floor((serverNow || Date.now()) / 1000) * 1000);
  const newRow = oldRow.slice();
  const manual = plan.manual || [];
  const leave = plan.leave || [];
  if (plan.clearAllMarks) Object.keys(meta.f).forEach(function(f) { delete meta.f[f].m; delete meta.f[f].L; });
  if (plan.clearLeaveMarks) Object.keys(meta.f).forEach(function(f) { delete meta.f[f].L; });
  fields.forEach(function(f) {
    newRow[COL[f] - 1] = writes[f];
    const prev = meta.f[f] || {};
    const entry = { t: at };
    if (seq) entry.q = seq; // 受付順（比較時刻が全く同じときの決め手。3.9-2）
    // 有給で空にする項目は、手修正印を外して「有給による空欄」印に置き換える（3.9-7）
    if (manual.indexOf(f) !== -1 || (prev.m && !plan.clearAllMarks && leave.indexOf(f) === -1)) entry.m = 1;
    if (leave.indexOf(f) !== -1) entry.L = 1;
    meta.f[f] = entry;
  });
  meta.savedAt = now.getTime(); // 新しい版の GAS が保存したサーバー時刻（10列目と同じ。3.11-1）
  newRow[COL.meta - 1] = JSON.stringify(meta);
  newRow[COL.serverUpdatedAt - 1] = now;
  // 11列目（更新時刻）は「今の値と操作時刻の大きい方」で、戻らない（3.6-1）
  const prevMs = (oldRow[COL.clientUpdatedAt - 1] && !isNaN(new Date(String(oldRow[COL.clientUpdatedAt - 1])).getTime()))
    ? new Date(String(oldRow[COL.clientUpdatedAt - 1])).getTime() : 0;
  newRow[COL.clientUpdatedAt - 1] = prevMs > at ? String(oldRow[COL.clientUpdatedAt - 1]) : new Date(at).toISOString();
  newRow[COL.sender - 1] = sender || '';
  // 備考の「有給申請」だけで有給と判定されていた既存行は、初めて新しい形式で書くときに8列目へ「有給」を書いて引き継ぐ（3.8-9）
  if (!hasMetaCell(oldRow) && rowIsPaidLeave(oldRow) && fields.indexOf('isPaidLeave') === -1) newRow[COL.isPaidLeave - 1] = '有給';
  if (plan.selfEdit) newRow[COL.selfEdited - 1] = '本人修正';
  if (plan.clearAllMarks) newRow[COL.selfEdited - 1] = '';
  const targetRow = exists ? found.rowIndex : Math.max(sheet.getLastRow() + 1, 2);
  // 1日分の1〜17列を1回で書き、同じ手順で EditLog も書いてから確定させる（3.9-2・3.10-1・3.11-4）。
  // EditLog の各行には opId を入れる。
  sheet.getRange(targetRow, 1, 1, SHEET_COLS).setValues([newRow]);
  if (!ctx.editLogged) {
    // ===== 修正履歴（EditLog）: 変更前の値（OpLog「受付」に残したもの）と書いた値を比べ、変わった項目だけ記録 =====
    // 【コードレビュー1-1】EditLog が書けなかったときは「EditLog 済み」の印を付けず、「受付」のまま残す（error）。
    // 再送で、書けている勤怠の行はそのままに EditLog だけ作り直す。
    try {
      writeEditLogForOp(ss, sheetName, staffName, dateStr, newRow, before, writes, fields, clockFix, opData.field, op.opId, sender, !!ctx.resumed);
      ctx.markEditLogged();
    } catch (e) {
      console.warn('EditLog failed:', e);
      SpreadsheetApp.flush();
      return { result: 'error', reason: 'editlog_failed', wrote: true };
    }
  }
  SpreadsheetApp.flush();
  return { result: 'applied', reason: plan.reason, corrected: correctedOut || (clockFix ? { written: false, field: opData.field } : undefined), wrote: true };
}

// 操作1つ分の EditLog（変更前の値 before.v → 書いた値 after）。失敗したら例外を投げる（呼ぶ側で印を付けない）
// 【コードレビュー2-1】当て直し（resumed）のときは、同じ opId・同じ内容の行が EditLog にもうあれば書かない
// （EditLog を書いた後、「EditLog 済み」の印を付ける前に止まっても二重にならない）。
function writeEditLogForOp(ss, sheetName, staffName, dateStr, baseRow, before, after, fields, clockFix, fixField, opId, sender, resumed) {
  const beforeRow = baseRow.slice();
  const afterRow = baseRow.slice();
  Object.keys(before.v || {}).forEach(function(f) { beforeRow[COL[f] - 1] = before.v[f]; });
  fields.forEach(function(f) { afterRow[COL[f] - 1] = after[f]; });
  const lines = [];
  if (before.exists) {
    const text = editChangesText(beforeRow, afterRow, fields);
    if (text) lines.push({ who: sheetName.replace(/_\d{6}$/, ''), date: fmtDateCell(baseRow[0], Session.getScriptTimeZone()) || dateStr, text: text });
  }
  // 時計補正の記録は、実際に時刻を書いたときだけ
  if (clockFix && fixField && fields.indexOf(fixField) !== -1 && (before.v || {})[fixField] !== after[fixField]) {
    lines.push({ who: staffName, date: dateStr, text: '時計補正 ' + clockFix.from + ' → ' + clockFix.to + '（送信元 ' + (sender || '不明') + '）' });
  }
  if (lines.length === 0) return;
  const already = resumed ? editLogTextsForOp(ss, opId) : [];
  lines.forEach(function(l) {
    if (already.indexOf(l.text) !== -1) return;
    appendEditLogLine(ss, l.who, l.date, l.text, opId);
  });
}

// EditLog の5列目（opId）が同じ行の「変更内容」を返す。TextFinder で opId の列だけを探す（無ければ直近の行を読む）
const EDITLOG_DEDUP_SCAN_ROWS = 2000;
function editLogTextsForOp(ss, opId) {
  if (!opId) return [];
  const log = ss.getSheetByName('EditLog');
  if (!log || log.getLastRow() < 2) return [];
  const last = log.getLastRow();
  const col = log.getRange(2, 5, last - 1, 1);
  const out = [];
  if (typeof col.createTextFinder === 'function') {
    col.createTextFinder(opId).matchEntireCell(true).findAll().forEach(function(r) { out.push(String(log.getRange(r.getRow(), 4).getValue() || '')); });
    return out;
  }
  const from = Math.max(2, last - EDITLOG_DEDUP_SCAN_ROWS + 1);
  log.getRange(from, 4, last - from + 1, 2).getValues().forEach(function(r) { if (String(r[1]) === opId) out.push(String(r[0] || '')); });
  return out;
}

// 書き込みの順番の鎖から、この操作の書き込みが届いていたかを決める（コードレビュー2-2・3-1）。
// 17列目のその項目の受付順（今の値を書いた操作）から始めて、その操作が書いたときに残した「直前に書いた操作の受付順」（pq）を
// さかのぼる。途中でこの操作の受付順に行き着けば、この操作はその項目を書いていた（後から上書きされた）。
// 受付順の前後には頼らない（先に受け付けて止まった操作が、後で再送されて上書きした場合もたどれる）。
// 【コードレビュー4】鎖でたどる操作の記録は、ロックの中でシートから読み直してから使う
// （同じまとめ送りの途中で、別の実行がその操作を当て直して re を足しているかもしれないため）。
function landedByWriteChain(ss, opLog, seq, stored, meta) {
  if (!opLog || !opLog.map || !stored || !stored.w || !(seq > 0)) return false;
  const bySeq = {};
  Object.keys(opLog.map).forEach(function(id) { const e = opLog.map[id]; if (e && Number(e.seq) > 0) bySeq[Number(e.seq)] = e; });
  return Object.keys(stored.w).some(function(f) {
    let q = (meta.f[f] && Number(meta.f[f].q)) || 0;
    const seen = {};
    while (q > 0 && !seen[q]) {
      if (q === seq) return true;
      seen[q] = true;
      let e = bySeq[q];
      if (e && e.ref) {
        const sh = ss.getSheetByName(e.ref.sheet);
        if (sh) e = opLogEntryFromRow(sh.getRange(e.ref.row, 1, 1, OPLOG_COLS).getValues()[0], e.ref.sheet, e.ref.row);
      }
      const plan = e && e.before ? effectivePlan(e.before) : null;
      if (!plan || !plan.pq || !Object.prototype.hasOwnProperty.call(plan.pq, f)) return false;
      q = Number(plan.pq[f]) || 0;
    }
    return false;
  });
}
// 最後に実際に使った計画（当て直しで決め直していれば re、無ければ最初の値）
function effectivePlan(b) {
  return b && b.re ? b.re : b;
}
// 返事に入れる時刻補正（最後に実際に使った計画のもの）
function finalCorrection(b) {
  const p = effectivePlan(b);
  return p && p.c ? p.c : undefined;
}

// 行の今の状態（ルール判断用）
// 新旧の比べ方: 直した操作時刻が大きい方が新しい。全く同じなら受付順（q）が後の方が新しい（3.6-1・3.9-2）。
// 受付順が記録されていない古い印とは「以上」で比べる（後から当てる方を新しいとする）。
function rowState(row, meta, seq) {
  const st = {
    deleted: row[COL.deleted - 1] === 'true' || row[COL.deleted - 1] === true,
    paidLeave: rowIsPaidLeave(row),
    info: function(f) { return fieldInfo(meta, f); },
    value: function(f) { return row[COL[f] - 1]; },
    empty: {},
    maxT: 0,
    maxQ: 0,
    seq: Number(seq) || 0
  };
  // fi = { t, q }。操作（at と この操作の受付順）が fi より古くなければ true
  st.notOlder = function(at, fi) {
    if (at !== fi.t) return at > fi.t;
    return !fi.q || !st.seq || st.seq >= fi.q;
  };
  st.notOlderThanField = function(at, f) { return st.notOlder(at, st.info(f)); };
  st.notOlderThanRow = function(at) { return st.notOlder(at, { t: st.maxT, q: st.maxQ }); };
  ALL_FIELDS.forEach(function(f) {
    const v = row[COL[f] - 1];
    st.empty[f] = (v === '' || v === null || v === undefined || (f === 'meal' && v !== '有' && v !== 1 && v !== true) || (f === 'deleted' && v !== 'true' && v !== true));
    const fi = fieldInfo(meta, f);
    if (fi.t > st.maxT || (fi.t === st.maxT && (fi.q || 0) > st.maxQ)) { st.maxT = fi.t; st.maxQ = fi.q || 0; }
  });
  return st;
}

// セルに書く値へ変換
function encodeField(f, v) {
  if (TIME_FIELDS.indexOf(f) !== -1) return normalizeTimeValue(v);
  if (f === 'meal') return v ? '有' : '';
  if (f === 'isPaidLeave') return v ? '有給' : '';
  if (f === 'deleted') return v ? 'true' : '';
  if (f === 'remarks') return (v === null || v === undefined) ? '' : String(v);
  if (f === 'additionalBreakMins' || f === 'transport') {
    if (v === null || v === undefined || v === '') return '';
    const n = parseInt(v, 10);
    return isNaN(n) ? '' : n;
  }
  return v;
}

function normalizeTimeValue(v) {
  if (!v) return '';
  const m = String(v).match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?/);
  if (!m) return '';
  return String(m[1]).padStart(2, '0') + ':' + m[2] + ':' + (m[3] || '00');
}

// 時刻欄を打刻で埋めてよいか: 空・手修正済みでない・直した操作時刻がその欄の時刻以上（3.6-2・3.7-3）
function canFillTime(st, f, at) {
  const fi = st.info(f);
  return st.empty[f] && !fi.m && st.notOlder(at, fi);
}
// 賄い・交通費・備考・休憩分数を打刻で書いてよいか: 手修正済みでない、かつ 操作時刻 ≧ 項目の時刻
function canWriteAttr(st, f, at) {
  const fi = st.info(f);
  return !fi.m && st.notOlder(at, fi);
}

// ----- ルール: punch（打刻ボタン・「情報のみ保存」）3.6-2・3.7-3 -----
function planPunch(st, at, data) {
  if (st.deleted) return { writes: {}, reason: 'deleted' };
  if (st.paidLeave) return { writes: {}, reason: 'paid_leave' };
  const writes = {};
  if (data.field && TIME_FIELDS.indexOf(data.field) !== -1 && data.time) {
    if (canFillTime(st, data.field, at)) writes[data.field] = encodeField(data.field, data.time);
  }
  const attrs = data.attrs || {};
  ATTR_FIELDS.forEach(function(f) {
    if (!Object.prototype.hasOwnProperty.call(attrs, f)) return; // 触っていない項目は送られない＝消さない
    if (canWriteAttr(st, f, at)) writes[f] = encodeField(f, attrs[f]);
  });
  return { writes: writes, reason: Object.keys(writes).length ? undefined : 'no_change_allowed' };
}

// ----- ルール: edit（修正画面・管理者の修正・有給申請/取消）3.6-2・3.7-13・3.8-4 -----
// 変えた項目だけが届く。「操作時刻 ≧ その項目の時刻」の項目だけ書き、手修正済みの印を付ける（管理者の修正も同じ）。
// 削除より後の edit は削除を外して書く。
// 有給の設定・取消は「行全体」に対する操作: 操作時刻がその行のどの項目の時刻よりも新しいときだけ効く（古ければ skipped）。
//   設定: 有給＝あり にして、出勤・中抜け・退勤・休憩分数・交通費・賄いを空にし、それらに「有給による空欄」印（L）。
//   取消: 有給＝なし にして「有給による空欄」印を外す。
const LEAVE_BLANK_FIELDS = ['clockIn', 'breakStart', 'breakEnd', 'clockOut', 'additionalBreakMins', 'transport', 'meal'];
function planEdit(st, at, data) {
  const fields = data.fields || {};
  const hasLeave = Object.prototype.hasOwnProperty.call(fields, 'isPaidLeave');
  const writes = {};
  const manual = [];
  const leave = [];
  if (hasLeave) {
    if (!st.notOlderThanRow(at)) return { writes: {}, reason: 'older_than_row' };
    writes.isPaidLeave = encodeField('isPaidLeave', fields.isPaidLeave === true);
    manual.push('isPaidLeave');
    if (fields.isPaidLeave === true) {
      LEAVE_BLANK_FIELDS.forEach(function(f) { writes[f] = encodeField(f, null); leave.push(f); });
    }
  }
  Object.keys(fields).forEach(function(f) {
    if (EDIT_FIELDS.indexOf(f) === -1 || f === 'isPaidLeave') return;
    if (fields.isPaidLeave === true && LEAVE_BLANK_FIELDS.indexOf(f) !== -1) return; // 有給の日は空のまま
    if (!st.notOlderThanField(at, f)) return; // 古い修正が新しい修正を消さない
    writes[f] = encodeField(f, fields[f]);
    manual.push(f);
  });
  const any = Object.keys(writes).length > 0;
  if (st.deleted && any) { writes.deleted = ''; }
  return {
    writes: writes, manual: manual, leave: leave,
    clearLeaveMarks: hasLeave && fields.isPaidLeave !== true,
    selfEdit: any && !!data.selfEdit, reason: any ? undefined : 'older_than_field'
  };
}

// ----- ルール: delete（完全取消）3.7-2 -----
// 直した操作時刻が、その行のどの項目の時刻よりも新しい（同時刻なら後着）ときだけ効く。
// 効いたら行の中身を空にし（今までどおり行は残す）、手修正印はすべて外す。
function planDelete(st, at) {
  if (st.deleted) return { writes: {}, reason: 'already_deleted' };
  if (!st.notOlderThanRow(at)) return { writes: {}, reason: 'older_than_row' };
  const writes = { deleted: 'true' };
  TIME_FIELDS.concat(['meal', 'isPaidLeave', 'remarks', 'additionalBreakMins', 'transport']).forEach(function(f) { writes[f] = encodeField(f, null); });
  return { writes: writes, clearAllMarks: true };
}

// ----- ルール: legacy（古い版の未送信）3.6-6・3.7-9 -----
// 端末は「既定値でない値」（時刻欄＝値あり、賄い＝あり、交通費＝数値、備考＝空でない、休憩分数＝1以上、有給＝あり、削除＝あり）を data に、
// 既定値だった項目の名前を data.cleared に入れて送る。サーバーはクラウドの今の値と項目ごとに比べて当てる／当てないを決め、
// 当てなかった項目は理由と一緒に返す（アプリの管理画面に一覧が出る）。
function planLegacy(st, at, rec) {
  rec = rec || {};
  const report = [];
  if (rec.deleted === true) {
    const p = planDelete(st, at);
    if (!Object.keys(p.writes).length) report.push('削除: ' + (p.reason === 'older_than_row' ? 'クラウドの方が新しい' : '削除済み'));
    p.reason = report.join(' / ') || undefined;
    return p;
  }
  if (st.deleted) return { writes: {}, reason: '削除済みの日' };
  const writes = {};
  TIME_FIELDS.forEach(function(f) {
    if (!rec[f]) return;
    if (st.paidLeave) { report.push(f + ': 有給の日'); return; }
    if (!st.empty[f]) { if (normalizeTimeValue(rec[f]) !== formatTimeCell(st.value(f))) report.push(f + ': クラウドに別の時刻あり'); return; }
    if (st.info(f).m) { report.push(f + ': 手修正済み'); return; }
    if (!st.notOlderThanField(at, f)) { report.push(f + ': クラウドの方が新しい'); return; }
    writes[f] = encodeField(f, rec[f]);
  });
  ['meal', 'transport', 'remarks', 'additionalBreakMins', 'isPaidLeave'].forEach(function(f) {
    if (!Object.prototype.hasOwnProperty.call(rec, f)) return;
    const fi = st.info(f);
    if (fi.m) { report.push(f + ': 手修正済み'); return; }
    if (!st.notOlder(at, fi)) { report.push(f + ': クラウドの方が新しい'); return; }
    writes[f] = encodeField(f, rec[f]);
  });
  (Array.isArray(rec.cleared) ? rec.cleared : []).forEach(function(f) {
    if (ALL_FIELDS.indexOf(f) === -1 || st.empty[f]) return;
    report.push(f + ': 空・なし・0 にする修正は運べない（意図が不明）');
  });
  return { writes: writes, reason: report.length ? report.join(' / ') : undefined };
}

const OP_PLANNERS = { punch: planPunch, edit: planEdit, 'delete': planDelete, legacy: planLegacy };

function blankRow(dateStr) {
  const p = dateStr.split('-');
  const dateObj = new Date(parseInt(p[0]), parseInt(p[1]) - 1, parseInt(p[2]));
  const dowNames = ['日', '月', '火', '水', '木', '金', '土'];
  const row = [];
  for (let i = 0; i < SHEET_COLS; i++) row.push('');
  row[COL.date - 1] = dateStr;
  row[COL.dow - 1] = dowNames[dateObj.getDay()];
  return row;
}

// 同じ日付の行を探す（重複があれば更新時刻が最新の行）
function findRowForDate(sheet, dateStr, tz) {
  const lastRow = sheet.getLastRow();
  const out = { rowIndex: -1, row: null, lastRow: lastRow };
  if (lastRow < 2) return out;
  let rows = [];
  try {
    rows = readSheetRows(sheet);
  } catch (e) {
    console.warn('Find existing row error: ' + sheet.getName(), e);
    throw e;
  }
  let bestMs = -1;
  for (let i = 0; i < rows.length; i++) {
    if (cellToDateStr(rows[i][0], tz) !== dateStr) continue;
    const ms = rowCompareMs(rows[i]);
    if (out.rowIndex === -1 || ms > bestMs) {
      out.rowIndex = i + 2;
      out.row = rows[i];
      bestMs = ms;
    }
  }
  return out;
}

function createStaffSheet(ss, sheetName) {
  const sheet = ss.insertSheet(sheetName);
  sheet.getRange(1, 1, 1, SHEET_COLS).setValues([SHEET_HEADERS]);
  sheet.getRange(1, 1, 1, SHEET_COLS).setFontWeight('bold').setBackground('#e8f5e9');
  sheet.setFrozenRows(1);
  sheet.setTabColor('#4caf50');
  return sheet;
}

// 既存のシートの最大列数が17列未満の場合は自動的に17列に拡張する（3.1(3)）
function ensureSheetColumns(sheet) {
  const currentMaxCols = sheet.getMaxColumns();
  if (currentMaxCols < SHEET_COLS) {
    sheet.insertColumnsAfter(currentMaxCols, SHEET_COLS - currentMaxCols);
    // 足りないヘッダーを自動補完
    const headerRange = sheet.getRange(1, currentMaxCols + 1, 1, SHEET_COLS - currentMaxCols);
    headerRange.setValues([SHEET_HEADERS.slice(currentMaxCols)]);
    headerRange.setFontWeight('bold').setBackground('#e8f5e9');
  }
}

// ===== OpLog（3.6-4・3.7-5）: 処理した opId と結果を残す（180日分） =====
// 【3.9-2・3.10-1】書く順番: ①「受付」（opId・比べる時刻・受付順）→ flush → ② 勤怠の行 → flush → ③「完了」（結果・理由）→ flush。
// 「受付」のまま残った opId は、再送のときに記録した比較時刻・受付順でもう一度当てる（何度当てても同じ結果）。
// 【3.10-3】OpLog は月ごとのシート（OpLog_YYYYMM、受け付けた月）に書き、行は消さない。照会は直近7か月分のシートを見る。
// 180日より古い月のシートは毎晩の自動実行で丸ごと消す（今月・先月は消さない）。
const OPLOG_ACCEPTED = 'accepted';
const OPLOG_COLS = 13;
const OPLOG_LOOKUP_MONTHS = 7;
const OPLOG_SEQ_PROP = 'OPLOG_SEQ';
function opLogSheetName(ms) {
  return OPLOG_SHEET_NAME + '_' + Utilities.formatDate(new Date(ms), 'Asia/Tokyo', 'yyyyMM');
}
function getOpLogSheet(ss, name, create) {
  let sh = ss.getSheetByName(name);
  if (!sh && create) {
    sh = ss.insertSheet(name);
    sh.setTabColor('#9e9e9e');
    sh.getRange(1, 1, 1, OPLOG_COLS).setValues([['opId', '結果', '理由', '受信時刻', '操作時刻', '種類', 'スタッフ', '日付', '送信元', '比べた時刻（直した操作時刻）', '受付順', '変更前の値', 'EditLog済み']]);
    sh.getRange(1, 1, 1, OPLOG_COLS).setFontWeight('bold').setBackground('#eeeeee');
  }
  return sh;
}
// 直近7か月分のシート名（新しい順）
function recentOpLogSheetNames(nowMs) {
  const d = new Date(nowMs);
  const names = [];
  for (let i = 0; i < OPLOG_LOOKUP_MONTHS; i++) {
    names.push(opLogSheetName(new Date(d.getFullYear(), d.getMonth() - i, 15).getTime()));
  }
  return names;
}

// { map: { opId: {result, reason, normAt, seq, ref:{sheet,row}, before, editLogged} }, lastRows: { シート名: 読んだ最後の行 } }
function opLogEntryFromRow(r, name, rowNo) {
  let before = null;
  try { before = r[11] ? JSON.parse(String(r[11])) : null; } catch (e) { before = null; }
  return { result: String(r[1] || ''), reason: String(r[2] || ''), normAt: Date.parse(String(r[9] || '')), seq: Number(r[10]) || 0, ref: { sheet: name, row: rowNo },
    before: before, editLogged: String(r[12] || '') === '1', staff: String(r[6] || ''), date: String(r[7] || '') };
}
function readOpLogRowsInto(sh, name, fromRow, toRow, map) {
  if (toRow < fromRow) return;
  const values = sh.getRange(fromRow, 1, toRow - fromRow + 1, OPLOG_COLS).getValues();
  values.forEach(function(r, i) {
    const id = String(r[0] || '');
    if (!id || map[id]) return; // 同じ opId は最初の行
    map[id] = opLogEntryFromRow(r, name, fromRow + i);
  });
}
function loadOpLog(ss) {
  const out = { map: {}, lastRows: {} };
  // 古い月から読む（同じ opId は最初の行）
  recentOpLogSheetNames(Date.now()).reverse().forEach(function(name) {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    const last = sh.getLastRow();
    out.lastRows[name] = last;
    readOpLogRowsInto(sh, name, 2, last, out.map);
  });
  return out;
}
// ロックを取るたびに呼ぶ（コードレビュー1-11）: 初回は全部読み、2回目からは他の実行が足した行と、この opId の行だけ読み直す
function refreshOpLog(ss, state, opId) {
  if (!state.loaded) {
    const l = loadOpLog(ss);
    state.map = l.map; state.lastRows = l.lastRows; state.loaded = true;
    return;
  }
  recentOpLogSheetNames(Date.now()).reverse().forEach(function(name) {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    const last = sh.getLastRow();
    readOpLogRowsInto(sh, name, Math.max(state.lastRows[name] || 1, 1) + 1, last, state.map);
    state.lastRows[name] = last;
  });
  const e = state.map[opId];
  if (e && e.ref) {
    const sh = ss.getSheetByName(e.ref.sheet);
    if (sh) state.map[opId] = opLogEntryFromRow(sh.getRange(e.ref.row, 1, 1, OPLOG_COLS).getValues()[0], e.ref.sheet, e.ref.row);
  }
}

// 受付順の番号（通し番号）を1つ進める。ScriptProperties に置き、打刻と同じロックの中で呼ぶ（3.11-2）。
// OpLog の行番号は使わないので、月のシートが替わっても戻らない。
function nextOpSeq() {
  const props = PropertiesService.getScriptProperties();
  const next = (Number(props.getProperty(OPLOG_SEQ_PROP)) || 0) + 1;
  props.setProperty(OPLOG_SEQ_PROP, String(next));
  return next;
}

function opLogRowValues(op, result, reason, serverNow, sender, normAt, seq, before) {
  return [String(op.opId || ''), result, String(reason || '').substring(0, 500), new Date(serverNow).toISOString(), String(op.at || ''),
    String(op.type || ''), String(op.staffName || ''), String(op.date || ''), sender || '', new Date(normAt).toISOString(), String(seq || ''),
    before ? JSON.stringify(before) : '', ''];
}

// 今月の OpLog シートに1行足して確定させ、その場所を返す
function appendOpLogRow(ss, values, serverNow) {
  const name = opLogSheetName(serverNow);
  const sh = getOpLogSheet(ss, name, true);
  const row = Math.max(sh.getLastRow() + 1, 2);
  const range = sh.getRange(row, 1, 1, OPLOG_COLS);
  range.setNumberFormat('@'); // 自動変換（日付・数値）を防ぐ
  range.setValues([values]);
  SpreadsheetApp.flush();
  return { sheet: name, row: row };
}

// 「EditLog 済み」の印（3.12-1）: EditLog を書いた同じ手順で付ける（この後の flush で一緒に確定）
function markOpLogEditLogged(ss, ref) {
  getOpLogSheet(ss, ref.sheet, true).getRange(ref.row, 13).setValue('1');
}

// 「受付」の行を「完了」にする（結果・理由の2列だけ書く）
function finishOpLogRow(ss, ref, result, reason) {
  const sh = getOpLogSheet(ss, ref.sheet, true);
  sh.getRange(ref.row, 2, 1, 2).setValues([[result, String(reason || '').substring(0, 500)]]);
  SpreadsheetApp.flush();
}

// 180日より古い月の OpLog シートを丸ごと消す（今月・先月は消さない）。毎晩の時間主導トリガーで実行する。
// 打刻と同じ getScriptLock を最大10秒待ち、取れなければ翌晩に回す。
// 反映手順で installOpLogPruneTrigger() を1回だけ実行してトリガーを作る。
function pruneOpLogDaily() {
  const ss = getTimecardSs();
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { console.warn('pruneOpLogDaily: ロックが取れないため翌晩に回します'); return; }
  try {
    const now = new Date();
    const keep = {};
    keep[opLogSheetName(now.getTime())] = true;
    keep[opLogSheetName(new Date(now.getFullYear(), now.getMonth() - 1, 15).getTime())] = true;
    const cutoff = now.getTime() - OPLOG_KEEP_DAYS * 86400000;
    const re = new RegExp('^' + OPLOG_SHEET_NAME + '_(\\d{4})(\\d{2})$');
    ss.getSheets().forEach(function(sh) {
      const m = sh.getName().match(re);
      if (!m || keep[sh.getName()]) return;
      const monthEnd = new Date(parseInt(m[1], 10), parseInt(m[2], 10), 1).getTime(); // 翌月1日0時
      if (monthEnd <= cutoff) ss.deleteSheet(sh);
    });
  } finally {
    lock.releaseLock();
  }
}

function installOpLogPruneTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'pruneOpLogDaily') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('pruneOpLogDaily').timeBased().everyDays(1).atHour(3).create();
}

// 反映後の確認用: OpLog 掃除の自動実行が1つだけ設定されているかを実行ログに出す
function checkOpLogPruneTrigger() {
  const n = ScriptApp.getProjectTriggers().filter(function(t) { return t.getHandlerFunction() === 'pruneOpLogDaily'; }).length;
  console.log(n === 1 ? 'OK: OpLog の掃除（pruneOpLogDaily）は1つ設定されています' : 'NG: pruneOpLogDaily の設定が ' + n + ' 個です。installOpLogPruneTrigger を実行してください');
  return n;
}

// EditLog シート（5列目に opId。3.11-4）
function getEditLogSheet(ss) {
  let log = ss.getSheetByName('EditLog');
  if (!log) {
    log = ss.insertSheet('EditLog');
    log.setTabColor('#9e9e9e');
    log.appendRow(['修正日時', 'スタッフ', '対象日', '変更内容', 'opId']);
    log.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#eeeeee');
  } else if (String(log.getRange(1, 5).getValue() || '') === '') {
    log.getRange(1, 5).setValue('opId');
  }
  return log;
}
// EditLog に1行足す（時計のずれ・時計補正など）
function appendEditLogLine(ss, who, dateStr, text, opId) {
  const log = getEditLogSheet(ss);
  log.appendRow([Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss'), who, dateStr || '', text, opId || '']);
}

// 時計のずれを EditLog に残す（同じ送信元は1時間に1回まで）
function logClockSkew(ss, sender, skewMs) {
  try {
    const cache = CacheService.getScriptCache();
    const key = 'skew_logged_' + String(sender || 'unknown').replace(/[^\w\-]/g, '_');
    if (cache.get(key)) return;
    cache.put(key, '1', 3600);
    appendEditLogLine(ss, '(端末) ' + (sender || '不明'), '', '時計のずれ: サーバーより ' + Math.round(-skewMs / 1000) + ' 秒（端末の時計を確認してください）');
  } catch (e) { console.warn('logClockSkew failed', e); }
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
    // 端末側が付与した設定変更時刻（新旧判定用）。無い場合は空。
    if (row[0] === 'settingsUpdatedAt') {
      settings.settingsUpdatedAt = row[1] ? String(row[1]) : '';
    }
  });
  // 月次確認（スタッフが「今月分を確認しました」を押した日時）を同梱
  // 【2026-11-08】対象月が日付に化けた既存行も日本時間で「YYYY-MM」に直して拾う。
  // 同じ人・同じ月が複数あれば最後の日時。confirmationOps は予備確認用（3.5-8）。
  try {
    const conf = ss.getSheetByName('Confirmations');
    if (conf && conf.getLastRow() >= 2) {
      const rows = readConfirmationRows(conf);
      const map = {};
      const ops = {};
      rows.forEach(function(r) {
        if (!r.staff || !r.ym) return;
        const key = r.staff + '|' + r.ym;
        if (map[key] === undefined || String(r.at) >= String(map[key])) {
          map[key] = r.at;
          if (r.opId) ops[key] = r.opId;
        }
      });
      settings.confirmations = map;
      settings.confirmationOps = ops;
    } else {
      settings.confirmations = {};
      settings.confirmationOps = {};
    }
    // 【コードレビュー1-2】本人確認を読めたことの印。エクセル出力はこの印が無ければ止める（全員「未」で出さない）
    settings.confirmationsOk = true;
  } catch (e) {
    console.warn('Confirmations の読み取りに失敗', e);
    settings.confirmationsOk = false;
    settings.confirmationsError = String(e && e.message ? e.message : e);
  }
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
  // 端末が付与した設定変更時刻を保存（端末側の新旧判定に使う）。無ければサーバー時刻のISO文字列。
  rows.push(['settingsUpdatedAt', settings.settingsUpdatedAt || new Date().toISOString()]);
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
// 【2026-11-08】対象行を先に決め（既存行は日本時間で「YYYY-MM」に直して照合）、
// そのセルを「書式なしテキスト」にしてから値を書く（appendRow は使わない。日付への自動変換を防ぐ）。（3.3-2・3.5-8）
// 【3.10-5】確認日時は「直した操作時刻・受付順」の組が今の記録より新しいときだけ上書きする（5・6列目に比較用の値を残す）
function confirmMonth(ss, staffName, yearMonth, opId, atMs, seq) {
  if (!staffName || !/^\d{4}-\d{2}$/.test(String(yearMonth || ''))) {
    return { status: 'error', result: 'invalid', message: 'スタッフ名と対象月を指定してください' };
  }
  let sheet = ss.getSheetByName('Confirmations');
  if (!sheet) {
    sheet = ss.insertSheet('Confirmations');
    sheet.setTabColor('#3f51b5');
    sheet.getRange(1, 1, 1, 6).setValues([['スタッフ名', '対象月', '確認日時', 'opId', '比較時刻', '受付順']]);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#e8eaf6');
  }
  const atVal = atMs ? Number(atMs) : Date.now();
  const now = Utilities.formatDate(new Date(atVal), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
  const rows = readConfirmationRows(sheet);
  let targetRow = -1;
  let cur = null;
  for (let i = 0; i < rows.length; i++) {
    if (rows[i].staff === String(staffName) && rows[i].ym === yearMonth) {
      if (opId && rows[i].opId === opId) return { status: 'ok', result: 'applied', reason: 'duplicate', confirmedAt: rows[i].at };
      targetRow = i + 2; // 上書き（最後に押した日時が有効）。複数あれば最後の行
      cur = rows[i];
    }
  }
  if (cur) {
    // 今の記録の比較時刻（無い既存行は「確認日時」の分）より古い確認では上書きしない
    const curT = cur.cmpAt || (Date.parse(String(cur.at).replace(' ', 'T') + ':00+09:00') || 0);
    const curQ = cur.cmpSeq || 0;
    const newer = atVal > curT || (atVal === curT && (!curQ || !seq || seq >= curQ));
    if (!newer) return { status: 'ok', result: 'skipped', message: 'older_than_confirmation', confirmedAt: cur.at };
  }
  if (targetRow === -1) targetRow = Math.max(sheet.getLastRow() + 1, 2);
  const range = sheet.getRange(targetRow, 1, 1, 6);
  range.setNumberFormat('@');
  range.setValues([[String(staffName), String(yearMonth), now, String(opId || ''), String(atVal), String(seq || '')]]);
  SpreadsheetApp.flush();
  return { status: 'ok', result: 'applied', confirmedAt: now };
}

function readConfirmationRows(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const cols = Math.min(Math.max(sheet.getLastColumn(), 3), 6);
  const values = sheet.getRange(2, 1, lastRow - 1, cols).getValues();
  return values.map(function(r) {
    const ymCell = r[1];
    let ym = '';
    if (ymCell instanceof Date) ym = Utilities.formatDate(ymCell, 'Asia/Tokyo', 'yyyy-MM');
    else {
      const m = String(ymCell || '').match(/^(\d{4})[-\/](\d{1,2})/);
      ym = m ? m[1] + '-' + String(m[2]).padStart(2, '0') : String(ymCell || '');
    }
    const atCell = r[2];
    const at = atCell instanceof Date ? Utilities.formatDate(atCell, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm') : String(atCell || '');
    return { staff: String(r[0] || ''), ym: ym, at: at, opId: cols >= 4 ? String(r[3] || '') : '',
      cmpAt: cols >= 5 ? (Number(r[4]) || 0) : 0, cmpSeq: cols >= 6 ? (Number(r[5]) || 0) : 0 };
  });
}

// ===== 修正履歴（EditLog）: 既存行の値が変わったときだけ自動追記 =====
// 【2026-11-08】サーバーで実際に書いた項目だけを比べる。時刻は「時:分」にそろえて比べる（書式違いだけの記録を出さない）。
function logEditIfChanged(ss, sheetName, oldRow, newRow, writtenFields, opId) {
  const text = editChangesText(oldRow, newRow, writtenFields);
  if (!text) return;
  const tz = Session.getScriptTimeZone();
  const log = getEditLogSheet(ss);
  log.appendRow([
    Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss'),
    sheetName.replace(/_\d{6}$/, ''),
    fmtDateCell(oldRow[0], tz),
    text,
    opId || ''
  ]);
}
// 変わった項目の説明文（変わっていなければ ''）
function editChangesText(oldRow, newRow, writtenFields) {
  const tz = Session.getScriptTimeZone();
  const fmt = function(v, isTime) {
    if (v === null || v === undefined || v === '') return '';
    if (isTime) {
      if (v instanceof Date) return Utilities.formatDate(v, tz, 'HH:mm');
      if (typeof v === 'number' && v > 0 && v < 1) return (formatTimeCell(v) || '').substring(0, 5);
      const m = String(v).match(/^(\d{1,2}):(\d{2})/);
      if (m) return String(m[1]).padStart(2, '0') + ':' + m[2];
    }
    if (v instanceof Date) return Utilities.formatDate(v, tz, 'HH:mm');
    return String(v);
  };
  // 比較対象: 出勤(3) 中抜け開始(4) 中抜け終了(5) 退勤(6) 賄い(7) 有給(8) 備考(9) 削除(12) 休憩分数(13) 交通費(14)
  const labels = {
    clockIn: '出勤', breakStart: '中抜け開始', breakEnd: '中抜け終了', clockOut: '退勤',
    meal: '賄い', isPaidLeave: '有給', remarks: '備考', deleted: '削除', additionalBreakMins: '休憩分数', transport: '交通費'
  };
  const changes = [];
  Object.keys(labels).forEach(function(f) {
    if (writtenFields.indexOf(f) === -1) return;
    const isTime = TIME_FIELDS.indexOf(f) !== -1;
    const o = fmt(oldRow[COL[f] - 1], isTime);
    const n = fmt(newRow[COL[f] - 1], isTime);
    if (o !== n) changes.push(labels[f] + ': 「' + (o || 'なし') + '」→「' + (n || 'なし') + '」');
  });
  return changes.join(' / ');
}

function fmtDateCell(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  return String(v || '');
}

/**
 * 【戻すとき用・6章／3.6-11】反映前に控えたスプレッドシート（コピー）と今のシートを行ごとに比べ、
 * 「控えと今で、どれかの項目が変わった行」と「今だけにある行（反映後に増えた行）」をすべてログに一覧表示する（何も書き換えない）。
 * EditLog・OpLog と照らして、必要な行だけ手で戻すために使う。
 * 使い方: スクリプトプロパティ BACKUP_SS_ID に控えのIDを入れて、エディタからこの関数を実行 →「実行ログ」を見る。
 */
function listRowsChangedSinceBackup() {
  const backupId = PropertiesService.getScriptProperties().getProperty('BACKUP_SS_ID');
  if (!backupId) { console.log('スクリプトプロパティ BACKUP_SS_ID を設定してください'); return []; }
  const tz = Session.getScriptTimeZone();
  const cur = getTimecardSs();
  const bak = openSpreadsheetChecked(backupId, '控え');
  const fields = TIME_FIELDS.concat(['meal', 'isPaidLeave', 'remarks', 'deleted', 'additionalBreakMins', 'transport']);
  const out = [];
  bak.getSheets().forEach(function(bs) {
    const name = bs.getName();
    const nm = name.match(/^(.+)_\d{6}$/);
    if (!nm || nm[1] === OPLOG_SHEET_NAME) return;
    const cs = cur.getSheetByName(name);
    if (!cs) { out.push(name + ': シートが無くなっています'); return; }
    readSheetRows(bs).forEach(function(br) {
      const dateStr = cellToDateStr(br[0], tz);
      if (!dateStr) return;
      const found = findRowForDate(cs, dateStr, tz);
      if (found.rowIndex === -1) { out.push(name + ' ' + dateStr + ': 行が無くなっています'); return; }
      const a = rowToRecord(br, nm[1], dateStr, tz), b = rowToRecord(found.row, nm[1], dateStr, tz);
      const diffs = fields.filter(function(f) { return String(a[f]) !== String(b[f]); })
        .map(function(f) { return f + ' 控え「' + a[f] + '」→ 今「' + b[f] + '」'; });
      if (diffs.length) out.push(name + ' ' + dateStr + ': ' + diffs.join(' / '));
    });
  });
  // 【コードレビュー1-5】今だけにあるシート・行（反映後に増えた記録）も一覧に出す
  cur.getSheets().forEach(function(cs) {
    const name = cs.getName();
    const nm = name.match(/^(.+)_\d{6}$/);
    if (!nm || nm[1] === OPLOG_SHEET_NAME) return;
    const bs = bak.getSheetByName(name);
    const bakDates = {};
    if (bs) readSheetRows(bs).forEach(function(br) { const d = cellToDateStr(br[0], tz); if (d) bakDates[d] = true; });
    readSheetRows(cs).forEach(function(cr) {
      const dateStr = cellToDateStr(cr[0], tz);
      if (!dateStr || bakDates[dateStr]) return;
      const b = rowToRecord(cr, nm[1], dateStr, tz);
      const vals = fields.filter(function(f) { return b[f] !== null && b[f] !== '' && b[f] !== false; })
        .map(function(f) { return f + '「' + b[f] + '」'; });
      out.push(name + ' ' + dateStr + ': ' + (bs ? '行が増えています' : 'シートごと増えています') + (vals.length ? '（' + vals.join(' / ') + '）' : ''));
    });
  });
  out.forEach(function(l) { console.log(l); });
  console.log('合計 ' + out.length + ' 件');
  return out;
}

/**
 * 【コードレビュー1-10・反映手順書 5】本番の Apps Script で1回だけ実行する: この勤怠ファイルの ID を
 * スクリプト プロパティ PROD_TIMECARD_SS_IDS に登録する（コードを書き換えずに「試験のときに本番を開かない仕組み」を使えるようにする）。
 * 試験用のコピー（MODE=test、または名前に「コピー」「控え」「試験」を含む）では登録しない。
 */
function registerProductionIds() {
  const props = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = ss.getName();
  if (props.getProperty('MODE') === 'test' || /コピー|控え|試験|test/i.test(name)) {
    console.log('NG: 試験用のファイル（' + name + '）では登録しません。本番の勤怠スプレッドシートの Apps Script で実行してください');
    return null;
  }
  props.setProperty('PROD_TIMECARD_SS_IDS', ss.getId());
  TEST_MODE_MEMO = null;
  console.log('OK: 本番の勤怠ファイル「' + name + '」を登録しました');
  return { timecard: ss.getId() };
}

/**
 * 【コードレビュー1-10・反映手順書 0-2】本番に出す前の試験（コピーしたスプレッドシート「タイムカード_試験用」の Apps Script で実行）。
 * 実際の GAS で次を確かめ、実行ログに「合格／不合格」を1行ずつ出す。最後の行が「全体: 合格」なら本番に進んでよい。
 *  1. 今月の勤怠をすべて読める（読めないシートが無い）・設定（名簿・本人確認）を読める
 *  2. 打刻 → 同じ opId の再送は二重にならない → 完全取消（試験用の名前「試験_リリース確認」で行い、最後にシートを消す）
 * 本番のファイルでは動かない（試験モードでなければ止まる）。有給の管理簿・計算用のファイルは使わない（2026-11-08 オーナー決定）。
 */
function runReleaseCheck() {
  const lines = [];
  let ok = true;
  const check = function(cond, msg) { lines.push((cond ? '合格: ' : '不合格: ') + msg); if (!cond) ok = false; };
  const finish = function() {
    lines.push(ok ? '全体: 合格' : '全体: 不合格（本番に進まないでください）');
    lines.forEach(function(l) { console.log(l); });
    return { ok: ok, lines: lines };
  };
  if (!isTestMode()) { ok = false; lines.push('不合格: 試験モードではありません（スクリプト プロパティに MODE=test を登録した、コピーのスプレッドシートで実行してください）'); return finish(); }
  try {
    const ss = getTimecardSs();
    const tz = Session.getScriptTimeZone();
    // 1. 読み取り
    const ym = Utilities.formatDate(new Date(), tz, 'yyyyMM');
    const all = JSON.parse(getAllRecords(resolveRecordRange({ month: ym })).getContent());
    check(all.status === 'ok' && Array.isArray(all.records), '今月の勤怠を読める（' + (all.records ? all.records.length : 0) + '件）');
    check(!(all.unreadableSheets && all.unreadableSheets.length), '読めないシートが無い' + (all.unreadableSheets && all.unreadableSheets.length ? '（' + all.unreadableSheets.join('、') + '）' : ''));
    const st = getSettings(ss);
    check(!!(st && Array.isArray(st.staffList) && st.staffList.length > 0), 'スタッフ名簿を読める');
    check(!!(st && st.confirmationsOk === true), '本人確認を読める');
    // 1-2. 有給管理簿を読む（読むだけ。管理簿は書き換えない）
    const lv = getLeaveInfo(Utilities.formatDate(new Date(), tz, 'yyyy'));
    const lvCount = lv && lv.leaveInfo ? Object.keys(lv.leaveInfo).length : 0;
    check(!!(lv && lv.found && lvCount > 0), '有給管理簿を読める（' + (lv && lv.fileName ? lv.fileName : (lv && lv.reason) || '見つからない') + '、' + lvCount + '人）');
    // 2. 打刻の往復
    const staff = '試験_リリース確認';
    const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    const id = 'release-' + Utilities.getUuid();
    const nowIso = new Date().toISOString();
    const send = function(ops) { return JSON.parse(handleOps(ss, { action: 'ops', clientVersion: MIN_CLIENT_VERSION, sender: '試験', sentAt: Date.now(), ops: ops }).getContent()); };
    const p = { opId: id + '-in', type: 'punch', staffName: staff, date: today, at: nowIso, offsetAtOp: 0, data: { field: 'clockIn', time: Utilities.formatDate(new Date(), tz, 'HH:mm:ss'), attrs: {} } };
    const r1 = send([p]);
    check(r1.status === 'ok' && r1.results[p.opId].result === 'applied', '打刻が書ける');
    const r2 = send([p]);
    check(r2.results[p.opId].duplicate === true, '同じ opId の再送は二重にならない');
    const d = { opId: id + '-del', type: 'delete', staffName: staff, date: today, at: new Date(Date.now() + 1000).toISOString(), offsetAtOp: 0, data: {} };
    const r3 = send([d]);
    const rec = getOneRecord(ss, staff, today).record;
    check(r3.results[d.opId].result === 'applied' && rec && rec.deleted === true, '完全取消が書ける');
    const sh = ss.getSheetByName(staff + '_' + today.substring(0, 4) + today.substring(5, 7));
    if (sh) ss.deleteSheet(sh);
  } catch (e) {
    check(false, '途中でエラー: ' + (e && e.message ? e.message : e));
  }
  return finish();
}
