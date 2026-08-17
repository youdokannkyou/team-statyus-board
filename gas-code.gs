/**
 * 用度環境課共有シート ― GAS 完成版(全機能統合・最終設計)
 *
 * 対象シート(6つ):
 *   プロジェクト / タスク / 報告内容 / メンバー / 確認ログ / 確認状態
 *
 * 通信フォーマットは全処理で統一:
 *   成功時: { "success": true, ... }
 *   失敗時: { "success": false, "error": "エラー内容" }
 *
 * 【設置方法】
 * 1. 対象スプレッドシートの「拡張機能」→「Apps Script」を開く
 * 2. 既存コードを全て削除し、これに置き換えて保存
 * 3. 「デプロイ」→「デプロイを管理」→ 鉛筆アイコン → 新しいバージョン → デプロイ
 *    (URLは変わらないので index.html 側の GAS_URL はそのまま)
 *
 * 【前提】以下6シートが実在すること
 *   プロジェクト: A id / B name / C createdAt / D status / E overview
 *   タスク:       A id / B projectId / C title / D assignee / E status / F due / G memo / H updatedAt
 *   報告内容:     A id / B reportDate / C author / D content / E lastConfirmedBy / F lastConfirmedAt / G lastConfirmMemo / H updatedAt
 *   メンバー:     A id / B name / C status(有効/無効)
 *   確認ログ:     A logId / B reportId / C confirmedBy / D confirmedAt / E type(confirm/unconfirm/comment) / F memo   ※追記専用(履歴)
 *   確認状態:     A reportId / B confirmedBy / C status(確認済み/未確認) / D updatedAt                                ※上書き型(現在の状態)
 */

const SHEET_PROJECTS      = 'プロジェクト';
const SHEET_TASKS         = 'タスク';
const SHEET_REPORTS       = '報告内容';
const SHEET_MEMBERS       = 'メンバー';
const SHEET_LOGS          = '確認ログ';
const SHEET_CONFIRM_STATE = '確認状態';

/* =========================================================
   エントリポイント
   ========================================================= */
function doGet(e) {
  try {
    const action = e.parameter.action || 'listAll';
    switch (action) {
      case 'listAll':           return jsonOutput_(buildAllData_());
      case 'listProjects':      return jsonOutput_({ success: true, projects: listProjects_() });
      case 'listTasks':         return jsonOutput_({ success: true, tasks: listTasks_() });
      case 'listReports':       return jsonOutput_({ success: true, reports: listReports_() });
      case 'listMembers':       return jsonOutput_({ success: true, members: listMembers_() });
      case 'listConfirmLogs':   return jsonOutput_({ success: true, confirmLogs: listConfirmLogs_() });
      case 'listConfirmStates': return jsonOutput_({ success: true, confirmStates: listConfirmStates_() });
      default:
        return jsonOutput_({ success: false, error: '不明なaction: ' + action });
    }
  } catch (err) {
    return jsonOutput_({ success: false, error: String(err) });
  }
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOutput_({ success: false, error: 'JSONの解析に失敗しました' });
  }

  try {
    switch (body.action) {
      // プロジェクト
      case 'createProject':     return jsonOutput_(createProject_(body));
      case 'updateProject':     return jsonOutput_(updateProject_(body));
      case 'deleteProject':     return jsonOutput_(deleteProject_(body));
      // タスク
      case 'createTask':        return jsonOutput_(createTask_(body));
      case 'updateTask':        return jsonOutput_(updateTask_(body));
      case 'deleteTask':        return jsonOutput_(deleteTask_(body));
      // 報告事項
      case 'createReport':      return jsonOutput_(createReport_(body));
      case 'deleteReport':      return jsonOutput_(deleteReport_(body));
      case 'confirmReport':     return jsonOutput_(confirmReport_(body));
      case 'unconfirmReport':   return jsonOutput_(unconfirmReport_(body));
      case 'addReportComment':  return jsonOutput_(addReportComment_(body));
      // メンバー
      case 'createMember':      return jsonOutput_(createMember_(body));
      case 'updateMemberName':  return jsonOutput_(updateMemberName_(body));
      case 'deactivateMember':  return jsonOutput_(deactivateMember_(body));
      default:
        return jsonOutput_({ success: false, error: '不明なaction: ' + body.action });
    }
  } catch (err) {
    return jsonOutput_({ success: false, error: String(err) });
  }
}

/* =========================================================
   全件まとめて取得(初回読み込み用)
   ========================================================= */
function buildAllData_() {
  return {
    success: true,
    projects: listProjects_(),
    tasks: listTasks_(),
    reports: listReports_(),
    members: listMembers_(),
    confirmLogs: listConfirmLogs_(),
    confirmStates: listConfirmStates_()
  };
}

/* =========================================================
   プロジェクト
   ========================================================= */
function listProjects_() {
  const rows = getRows_(SHEET_PROJECTS);
  return rows.map(r => ({
    id: String(r[0]),
    name: r[1],
    createdAt: r[2],
    status: r[3],
    overview: r[4] || ''
  }));
}

function createProject_(body) {
  if (!body.name) return { success: false, error: 'プロジェクト名が空です' };
  const sheet = getSheet_(SHEET_PROJECTS);
  const id = body.id || Utilities.getUuid();
  const createdAt = body.createdAt || todayStr_();
  const status = body.status || '進行中';
  const overview = body.overview || '';
  sheet.appendRow([id, body.name, createdAt, status, overview]);
  return { success: true, project: { id: id, name: body.name, createdAt: createdAt, status: status, overview: overview } };
}

// プロジェクト名・概要の更新(作成日・状態は変更しない)
function updateProject_(body) {
  if (!body.id) return { success: false, error: 'idが指定されていません' };
  const sheet = getSheet_(SHEET_PROJECTS);
  const rowIndex = findRowIndexById_(sheet, 0, body.id);
  if (rowIndex === -1) return { success: false, error: '対象のプロジェクトが見つかりません' };
  sheet.getRange(rowIndex, 2, 1, 1).setValue(body.name || '');       // B列: 名前
  sheet.getRange(rowIndex, 5, 1, 1).setValue(body.overview || '');   // E列: 概要
  const row = sheet.getRange(rowIndex, 1, 1, 5).getValues()[0];
  return { success: true, project: { id: String(row[0]), name: row[1], createdAt: row[2], status: row[3], overview: row[4] || '' } };
}

function deleteProject_(body) {
  if (!body.id) return { success: false, error: 'idが指定されていません' };
  const sheet = getSheet_(SHEET_PROJECTS);
  const found = deleteRowById_(sheet, 0, body.id);
  if (!found) return { success: false, error: '対象のプロジェクトが見つかりません' };
  // 紐づくタスクも削除
  deleteRowsWhere_(getSheet_(SHEET_TASKS), 1, body.id);
  return { success: true, id: body.id };
}

/* =========================================================
   タスク
   ========================================================= */
function listTasks_() {
  const rows = getRows_(SHEET_TASKS);
  return rows.map(r => ({
    id: String(r[0]),
    projectId: String(r[1]),
    title: r[2],
    assignee: r[3],
    status: r[4],
    due: formatDateOnly_(r[5]),
    memo: r[6],
    updatedAt: r[7]
  }));
}

function createTask_(body) {
  const sheet = getSheet_(SHEET_TASKS);
  const id = body.id || Utilities.getUuid();
  const updatedAt = nowStr_();
  const title = body.title || '';
  const assignee = body.assignee || '';
  const status = body.status || 'todo';
  const due = body.due || '';
  const memo = body.memo || '';
  sheet.appendRow([id, body.projectId, title, assignee, status, due, memo, updatedAt]);
  return { success: true, task: { id: id, projectId: String(body.projectId), title: title, assignee: assignee, status: status, due: due, memo: memo, updatedAt: updatedAt } };
}

function updateTask_(body) {
  if (!body.id) return { success: false, error: 'idが指定されていません' };
  const sheet = getSheet_(SHEET_TASKS);
  const rowIndex = findRowIndexById_(sheet, 0, body.id);
  if (rowIndex === -1) return { success: false, error: '対象のタスクが見つかりません' };
  const updatedAt = nowStr_();
  const title = body.title || '';
  const assignee = body.assignee || '';
  const status = body.status || 'todo';
  const due = body.due || '';
  const memo = body.memo || '';
  // C〜H列(タイトル・担当・状態・期限・メモ・更新日時)を更新。A/B(id, projectId)は変更しない
  sheet.getRange(rowIndex, 3, 1, 6).setValues([[title, assignee, status, due, memo, updatedAt]]);
  const projectId = sheet.getRange(rowIndex, 2, 1, 1).getValue();
  return { success: true, task: { id: body.id, projectId: String(projectId), title: title, assignee: assignee, status: status, due: due, memo: memo, updatedAt: updatedAt } };
}

function deleteTask_(body) {
  if (!body.id) return { success: false, error: 'idが指定されていません' };
  const sheet = getSheet_(SHEET_TASKS);
  const found = deleteRowById_(sheet, 0, body.id);
  if (!found) return { success: false, error: '対象のタスクが見つかりません' };
  return { success: true, id: body.id };
}

/* =========================================================
   報告事項
   ========================================================= */
function listReports_() {
  const rows = getRows_(SHEET_REPORTS);
  return rows.map(r => ({
    id: String(r[0]),
    reportDate: r[1],
    author: r[2],
    content: r[3],
    lastConfirmedBy: r[4],
    lastConfirmedAt: r[5],
    lastConfirmMemo: r[6],
    updatedAt: r[7]
  }));
}

function createReport_(body) {
  if (!body.content) return { success: false, error: '報告内容が空です' };
  const sheet = getSheet_(SHEET_REPORTS);
  const id = body.id || Utilities.getUuid();
  const now = nowStr_();
  const author = body.author || '';
  // B列(報告日)には日時まで含めて保存し、投稿順の並び替えに使えるようにする
  sheet.appendRow([id, now, author, body.content, '', '', '', now]);
  // 投稿者は自動的に「確認済み」扱いとする(既存サイトの挙動を維持)
  const confirmState = setConfirmState_(id, author, '確認済み');
  const confirmLog = appendConfirmLog_(id, author, 'confirm', '');
  const report = updateReportLatest_(id, author, confirmState.updatedAt, '');
  return { success: true, report: report, confirmState: confirmState, confirmLog: confirmLog };
}

function deleteReport_(body) {
  if (!body.id) return { success: false, error: 'idが指定されていません' };
  const sheet = getSheet_(SHEET_REPORTS);
  const found = deleteRowById_(sheet, 0, body.id);
  if (!found) return { success: false, error: '対象の報告が見つかりません' };
  deleteRowsWhere_(getSheet_(SHEET_LOGS), 1, body.id);
  deleteRowsWhere_(getSheet_(SHEET_CONFIRM_STATE), 0, body.id);
  return { success: true, id: body.id };
}

// 既読チェック(状態を「確認済み」に)
function confirmReport_(body) {
  if (!body.reportId || !body.confirmedBy) {
    return { success: false, error: 'reportId / confirmedBy が必要です' };
  }
  const confirmState = setConfirmState_(body.reportId, body.confirmedBy, '確認済み');
  const confirmLog = appendConfirmLog_(body.reportId, body.confirmedBy, 'confirm', '');
  const report = updateReportLatest_(body.reportId, body.confirmedBy, confirmState.updatedAt, '');
  return { success: true, confirmState: confirmState, confirmLog: confirmLog, report: report };
}

// 既読の取り消し(状態を「未確認」に)
function unconfirmReport_(body) {
  if (!body.reportId || !body.confirmedBy) {
    return { success: false, error: 'reportId / confirmedBy が必要です' };
  }
  const confirmState = setConfirmState_(body.reportId, body.confirmedBy, '未確認');
  const confirmLog = appendConfirmLog_(body.reportId, body.confirmedBy, 'unconfirm', '');
  const report = updateReportLatest_(body.reportId, body.confirmedBy, confirmState.updatedAt, '');
  return { success: true, confirmState: confirmState, confirmLog: confirmLog, report: report };
}

// コメント送信(既読も兼ねる)
function addReportComment_(body) {
  if (!body.reportId || !body.author || !body.content) {
    return { success: false, error: 'reportId / author / content が必要です' };
  }
  const confirmState = setConfirmState_(body.reportId, body.author, '確認済み');
  const confirmLog = appendConfirmLog_(body.reportId, body.author, 'comment', body.content);
  const report = updateReportLatest_(body.reportId, body.author, confirmState.updatedAt, body.content);
  return { success: true, confirmState: confirmState, confirmLog: confirmLog, report: report };
}

// 報告内容シートのE〜H列(最新の確認者・確認日・確認メモ・更新日時)を更新し、
// 更新後の行全体をオブジェクトとして返す(confirm/unconfirm/comment のレスポンスに含めるため)
function updateReportLatest_(reportId, confirmedBy, confirmedAt, memo) {
  const sheet = getSheet_(SHEET_REPORTS);
  const rowIndex = findRowIndexById_(sheet, 0, reportId);
  if (rowIndex === -1) return null;
  sheet.getRange(rowIndex, 5, 1, 4).setValues([[confirmedBy, confirmedAt, memo || '', confirmedAt]]);
  const row = sheet.getRange(rowIndex, 1, 1, 8).getValues()[0];
  return {
    id: String(row[0]), reportDate: row[1], author: row[2], content: row[3],
    lastConfirmedBy: row[4], lastConfirmedAt: row[5], lastConfirmMemo: row[6], updatedAt: row[7]
  };
}

/* =========================================================
   確認ログ(履歴・追記専用)
   ========================================================= */
function listConfirmLogs_() {
  const rows = getRows_(SHEET_LOGS);
  return rows.map(r => ({
    id: String(r[0]),
    reportId: String(r[1]),
    confirmedBy: r[2],
    confirmedAt: r[3],
    type: r[4],   // confirm / unconfirm / comment
    memo: r[5]
  }));
}

function appendConfirmLog_(reportId, confirmedBy, type, memo) {
  const sheet = getSheet_(SHEET_LOGS);
  const id = Utilities.getUuid();
  const confirmedAt = nowStr_();
  sheet.appendRow([id, reportId, confirmedBy, confirmedAt, type, memo || '']);
  return { id: id, reportId: String(reportId), confirmedBy: confirmedBy, confirmedAt: confirmedAt, type: type, memo: memo || '' };
}

/* =========================================================
   確認状態(現在の確認済み/未確認・上書き型)
   ========================================================= */
function listConfirmStates_() {
  const rows = getRows_(SHEET_CONFIRM_STATE);
  return rows.map(r => ({
    reportId: String(r[0]),
    confirmedBy: r[1],
    status: r[2], // 確認済み / 未確認
    updatedAt: r[3]
  }));
}

// 報告ID + 確認者 の組み合わせで1行に定まるよう upsert する
function setConfirmState_(reportId, confirmedBy, status) {
  const sheet = getSheet_(SHEET_CONFIRM_STATE);
  const lastRow = sheet.getLastRow();
  const now = nowStr_();

  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (let i = 0; i < values.length; i++) {
      if (String(values[i][0]) === String(reportId) && String(values[i][1]) === String(confirmedBy)) {
        sheet.getRange(i + 2, 3, 1, 2).setValues([[status, now]]);
        return { reportId: String(reportId), confirmedBy: confirmedBy, status: status, updatedAt: now };
      }
    }
  }
  sheet.appendRow([reportId, confirmedBy, status, now]);
  return { reportId: String(reportId), confirmedBy: confirmedBy, status: status, updatedAt: now };
}

/* =========================================================
   メンバー
   ========================================================= */
function listMembers_() {
  const rows = getRows_(SHEET_MEMBERS);
  return rows.map(r => ({
    id: String(r[0]),
    name: r[1],
    status: r[2]
  }));
}

function createMember_(body) {
  if (!body.name) return { success: false, error: 'メンバー名が空です' };
  const sheet = getSheet_(SHEET_MEMBERS);
  const id = nextMemberId_(sheet);
  sheet.appendRow([id, body.name, '有効']);
  return { success: true, member: { id: id, name: body.name, status: '有効' } };
}

function updateMemberName_(body) {
  if (!body.id || !body.name) return { success: false, error: 'id / name が必要です' };
  const sheet = getSheet_(SHEET_MEMBERS);
  const rowIndex = findRowIndexById_(sheet, 0, body.id);
  if (rowIndex === -1) return { success: false, error: '対象のメンバーが見つかりません' };
  sheet.getRange(rowIndex, 2, 1, 1).setValue(body.name);
  const status = sheet.getRange(rowIndex, 3, 1, 1).getValue();
  return { success: true, member: { id: body.id, name: body.name, status: status } };
}

function deactivateMember_(body) {
  if (!body.id) return { success: false, error: 'idが指定されていません' };
  const sheet = getSheet_(SHEET_MEMBERS);
  const rowIndex = findRowIndexById_(sheet, 0, body.id);
  if (rowIndex === -1) return { success: false, error: '対象のメンバーが見つかりません' };
  sheet.getRange(rowIndex, 3, 1, 1).setValue('無効');
  const name = sheet.getRange(rowIndex, 2, 1, 1).getValue();
  return { success: true, member: { id: body.id, name: name, status: '無効' } };
}

function nextMemberId_(sheet) {
  const rows = getRowsFromSheet_(sheet);
  let maxNum = 0;
  rows.forEach(r => {
    const m = String(r[0]).match(/^M(\d+)$/);
    if (m) maxNum = Math.max(maxNum, parseInt(m[1], 10));
  });
  const next = maxNum + 1;
  return 'M' + String(next).padStart(3, '0');
}

/* =========================================================
   共通ヘルパー
   ========================================================= */
function getSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('「' + name + '」シートが見つかりません');
  return sheet;
}

// ヘッダー行(1行目)を除いた全データ行を返す
function getRows_(sheetName) {
  return getRowsFromSheet_(getSheet_(sheetName));
}
function getRowsFromSheet_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  return values.filter(r => r[0] !== '' && r[0] !== null);
}

// colIndex は0始まり(A列=0)。見つかった場合はスプレッドシート上の行番号(1始まり)を返す。無ければ-1
function findRowIndexById_(sheet, colIndex, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const values = sheet.getRange(2, colIndex + 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

function deleteRowById_(sheet, colIndex, id) {
  const rowIndex = findRowIndexById_(sheet, colIndex, id);
  if (rowIndex === -1) return false;
  sheet.deleteRow(rowIndex);
  return true;
}

// colIndexに一致する行を全て削除(下から上へ処理してインデックスずれを防ぐ)
function deleteRowsWhere_(sheet, colIndex, value) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (String(values[i][colIndex]) === String(value)) {
      sheet.deleteRow(i + 2);
    }
  }
}

function nowStr_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}
function todayStr_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
function formatDateOnly_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return value ? String(value) : '';
}

function jsonOutput_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
