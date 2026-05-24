/**
 * 【領収書・請求書 AI抽出・管理アプリ (MVP版) WebApp.gs】
 *
 * Drive上の領収書・請求書をGemini APIで読み取り、確認後に
 * スプレッドシートへ記録し、処理済み/エラー用フォルダへ移動します。
 * 運用担当者が後から直しやすいよう、設定値はScript Propertiesで管理します。
 */

// 高速・低コストで画像/PDF入力に対応する現行モデル。
// Script Propertiesに GEMINI_MODEL を設定すると、環境ごとにモデルを差し替えられます。
const DEFAULT_GEMINI_MODEL = 'gemini-2.5-flash-lite';

// スプレッドシートのシート名と列定義を一箇所に集約し、保存処理のずれを防ぎます。
const RECEIPTS_SHEET_NAME = 'receipts';
const USERS_SHEET_NAME = 'users';
const RECEIPTS_HEADERS = ['保存日時', '元のファイル名', 'リネーム後ファイル名', '利用者名', '日付', '摘要', '金額', 'ファイルURL', 'ステータス'];

function doGet() {
  return HtmlService.createTemplateFromFile('Prototype')
    .evaluate()
    .setTitle('領収書・請求書 AI抽出アプリ')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getInputFilesFromDrive() {
  const folderId = PropertiesService.getScriptProperties().getProperty('RECEIPT_INPUT_FOLDER_ID');
  if (!folderId) return [];
  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    const fileList = [];
    while (files.hasNext()) {
      const file = files.next();
      const mimeType = file.getMimeType();
      if (mimeType === 'application/pdf' || mimeType === 'image/jpeg' || mimeType === 'image/png') {
        fileList.push({
          id: file.getId(),
          // 画面側・保存側で共通して扱えるよう、DriveファイルID専用の名前も付けます。
          driveFileId: file.getId(),
          name: file.getName(),
          fileName: file.getName(),
          mimeType: mimeType,
          size: file.getSize(),
          createdTime: file.getDateCreated().toISOString()
        });
      }
    }
    return fileList;
  } catch (e) {
    throw new Error('処理前フォルダの取得に失敗しました: ' + e.message);
  }
}

function getFileBase64(fileId) {
  if (!fileId) throw new Error('DriveファイルIDが指定されていません。');
  try {
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    const base64 = Utilities.base64Encode(blob.getBytes());
    return `data:${file.getMimeType()};base64,${base64}`;
  } catch (e) {
    throw new Error('ファイルの取得に失敗しました: ' + e.message);
  }
}

function analyzeReceiptPrototypeUpload(upload) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('GEMINI_API_KEYが設定されていません。');
  if (!upload) throw new Error('解析対象ファイルが指定されていません。');

  // PCアップロードはbase64を、Drive同期ファイルはDrive IDを使います。
  const driveFileId = upload.driveFileId || upload.id;
  let base64Data = normalizeBase64(upload.base64);
  if (!base64Data && driveFileId) {
    const dataUrl = getFileBase64(driveFileId);
    base64Data = dataUrl.split(',')[1];
  }
  if (!base64Data) throw new Error('解析対象ファイルのデータを取得できませんでした。');

  const mimeType = upload.mimeType || 'application/pdf';
  const model = getGeminiModel();

  const userMaster = getUsersMaster();
  const userMasterPrompt = userMaster.length > 0 
    ? `また、以下の利用者リスト（正式名称）と照合し、最も可能性が高いものをuserNameの候補に入れてください: [${userMaster.join(', ')}]`
    : "";

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

  const prompt = [
    '領収書または請求書の画像/PDFから、以下の項目を抽出してください。',
    '対象は日本語書類が中心です。請求書では宛名、領収書では利用者・宛名・支払者に近い氏名を優先してください。',
    '人が手動で確認・修正できるように、推測できる候補値を確信度が高い順に複数返してください。',
    '日付はYYYY-MM-DD、金額は税込合計を半角数字のみで返してください。',
    userMasterPrompt,
    '出力は必ずマークダウンを含まない、純粋なJSONオブジェクトのみとしてください。',
    '',
    '【JSON スキーマ】',
    '{',
    '  "userName": {"candidates": ["候補1 (利用者氏名)", "候補2"]},',
    '  "date": {"candidates": ["YYYY-MM-DD", "YYYY-MM-DD"]},',
    '  "summary": {"candidates": ["摘要・品目1", "摘要・品目2"]},',
    '  "amount": {"candidates": ["半角数字のみ金額1", "半角数字のみ金額2"]}',
    '}'
  ].join('\n');

  const payload = {
    contents: [{
      parts: [
        { inlineData: { mimeType: mimeType, data: base64Data } },
        { text: prompt }
      ]
    }],
    generationConfig: {
      responseMimeType: 'application/json'
    }
  };

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error('Gemini APIエラー: ' + response.getContentText());
  }

  const result = JSON.parse(response.getContentText());
  const rawText = result.candidates && result.candidates[0] &&
    result.candidates[0].content && result.candidates[0].content.parts &&
    result.candidates[0].content.parts[0] && result.candidates[0].content.parts[0].text;
  if (!rawText) throw new Error('Gemini APIの応答から抽出結果を取得できませんでした。');
  
  return {
    fields: normalizeGeminiFields(parseGeminiJson(rawText)),
    model: model
  };
}

function saveReceiptPrototypeReview(payload) {
  if (!payload.review) throw new Error('確認データがありません');
  const props = PropertiesService.getScriptProperties().getProperties();
  const spreadsheetId = props.RECEIPT_PROTOTYPE_SPREADSHEET_ID;
  const processedFolderId = props.RECEIPT_PROCESSED_FOLDER_ID;
  const errorFolderId = props.RECEIPT_ERROR_FOLDER_ID;

  const originalFileName = payload.fileName || payload.name || 'uploaded_file';
  const extension = getFileExtension(originalFileName, payload.mimeType);
  const userName = sanitizeString(payload.review.userName || 'unknown');
  const date = sanitizeString(payload.review.date || 'unknown');
  const summary = sanitizeString(payload.review.summary || '領収書');
  const storedFileName = `${userName}_${date}_${summary}${extension}`;

  let driveFileUrl = '';
  let targetFolderId = (payload.status === 'error' || payload.review.isError) ? errorFolderId : processedFolderId;
  const driveFileId = payload.driveFileId || payload.id;

  // Drive同期ファイルは元ファイルをリネームして指定フォルダへ移動します。
  if (driveFileId && targetFolderId) {
    try {
      const file = DriveApp.getFileById(driveFileId);
      file.setName(storedFileName);
      file.moveTo(DriveApp.getFolderById(targetFolderId));
      driveFileUrl = file.getUrl();
    } catch (e) {
      throw new Error('Driveファイルのリネーム・移動に失敗しました: ' + e.message);
    }
  } else if (targetFolderId && payload.base64) {
    // PCアップロードファイルはGAS側で新規ファイルとして保存します。
    try {
      const blob = Utilities.newBlob(Utilities.base64Decode(normalizeBase64(payload.base64)), payload.mimeType, storedFileName);
      const file = DriveApp.getFolderById(targetFolderId).createFile(blob);
      driveFileUrl = file.getUrl();
    } catch (e) {
      throw new Error('アップロードファイルの保存に失敗しました: ' + e.message);
    }
  }

  if (!spreadsheetId) throw new Error('RECEIPT_PROTOTYPE_SPREADSHEET_IDが設定されていません。');
  if (spreadsheetId) {
    try {
      const ss = SpreadsheetApp.openById(spreadsheetId);
      const sheet = ensureReceiptsSheet(ss);
      sheet.appendRow([new Date(), originalFileName, storedFileName, payload.review.userName, payload.review.date, payload.review.summary, payload.review.amount, driveFileUrl, payload.status]);
    } catch (e) {
      throw new Error('スプレッドシートへの保存に失敗しました: ' + e.message);
    }
  }
  return { storedFileName: storedFileName, driveFileUrl: driveFileUrl };
}

function getUsersMaster() {
  const ssId = PropertiesService.getScriptProperties().getProperty('RECEIPT_PROTOTYPE_SPREADSHEET_ID');
  if (!ssId) return [];
  try {
    const sheet = SpreadsheetApp.openById(ssId).getSheetByName(USERS_SHEET_NAME);
    if (!sheet) return [];
    const values = sheet.getDataRange().getValues();
    return values.flat().filter(v => v && v !== '氏名');
  } catch (e) { return []; }
}

function addOrDeleteUserMaster(action, name) {
  const ssId = PropertiesService.getScriptProperties().getProperty('RECEIPT_PROTOTYPE_SPREADSHEET_ID');
  if (!ssId) throw new Error('利用者マスタ用のスプレッドシートが未設定です。先に環境設定を実行してください。');
  if (!name) throw new Error('利用者名が空です。');
  const ss = SpreadsheetApp.openById(ssId);
  let sheet = ss.getSheetByName(USERS_SHEET_NAME) || ss.insertSheet(USERS_SHEET_NAME);
  if (action === 'add') {
    sheet.appendRow([name]);
  } else if (action === 'delete') {
    const values = sheet.getDataRange().getValues().flat();
    const idx = values.indexOf(name);
    if (idx !== -1) sheet.deleteRow(idx + 1);
  }
  return getUsersMaster();
}

function getDashboardStats() {
  const props = PropertiesService.getScriptProperties().getProperties();
  const stats = { folderStats: { input: 0, processed: 0, error: 0 }, totalSaved: 0, history: [], spreadsheetUrl: '' };
  
  const countFiles = (id) => {
    if (!id) return 0;
    try {
      let count = 0;
      const files = DriveApp.getFolderById(id).getFiles();
      while(files.hasNext()) { files.next(); count++; }
      return count;
    } catch(e) { return 0; }
  };

  stats.folderStats.input = countFiles(props.RECEIPT_INPUT_FOLDER_ID);
  stats.folderStats.processed = countFiles(props.RECEIPT_PROCESSED_FOLDER_ID);
  stats.folderStats.error = countFiles(props.RECEIPT_ERROR_FOLDER_ID);

  if (props.RECEIPT_PROTOTYPE_SPREADSHEET_ID) {
    try {
      const ss = SpreadsheetApp.openById(props.RECEIPT_PROTOTYPE_SPREADSHEET_ID);
      stats.spreadsheetUrl = ss.getUrl();
      const sheet = ensureReceiptsSheet(ss);
      const lastRow = sheet.getLastRow();
      stats.totalSaved = Math.max(0, lastRow - 1);
      if (lastRow > 1) {
        const startRow = Math.max(2, lastRow - 14);
        const numRows = lastRow - startRow + 1;
        const data = sheet.getRange(startRow, 1, numRows, 9).getValues();
        stats.history = data.reverse().map(row => ({
          savedAt: row[0] instanceof Date ? Utilities.formatDate(row[0], "JST", "yyyy-MM-dd HH:mm:ss") : row[0],
          userName: row[3],
          date: row[4] instanceof Date ? Utilities.formatDate(row[4], "JST", "yyyy-MM-dd") : row[4],
          summary: row[5],
          amount: row[6],
          driveFileUrl: row[7],
          status: row[8] === 'error' ? 'error' : 'processed'
        }));
      }
    } catch(e) {}
  }
  return stats;
}

function setupPrototypeEnvironment() {
  try {
    const root = DriveApp.createFolder('領収書AI抽出システム（デモ用）');
    const input = root.createFolder('01_処理前 (Input)');
    const processed = root.createFolder('02_処理済み (Processed)');
    const error = root.createFolder('03_エラー (Error)');
    const ss = SpreadsheetApp.create('領収書データ管理簿');
    DriveApp.getFileById(ss.getId()).moveTo(root);
    const sheet = ss.getSheets()[0];
    sheet.setName(RECEIPTS_SHEET_NAME);
    sheet.appendRow(RECEIPTS_HEADERS);
    const userSheet = ss.insertSheet(USERS_SHEET_NAME);
    userSheet.appendRow(['氏名']);
    userSheet.appendRow(['テスト 太郎']);
    const props = {
      RECEIPT_INPUT_FOLDER_ID: input.getId(),
      RECEIPT_PROCESSED_FOLDER_ID: processed.getId(),
      RECEIPT_ERROR_FOLDER_ID: error.getId(),
      RECEIPT_PROTOTYPE_SPREADSHEET_ID: ss.getId()
    };
    PropertiesService.getScriptProperties().setProperties(props);
    return { spreadsheetId: ss.getId(), inputFolderId: input.getId() };
  } catch (e) { throw new Error('失敗: ' + e.message); }
}

function sanitizeString(str) {
  return String(str).trim().replace(/[\\/:*?"<>|#%{}~&]/g, '_').replace(/\s+/g, '_').slice(0, 30);
}

function getGeminiModel() {
  return PropertiesService.getScriptProperties().getProperty('GEMINI_MODEL') || DEFAULT_GEMINI_MODEL;
}

function normalizeBase64(value) {
  if (!value) return '';
  return String(value).indexOf(',') >= 0 ? String(value).split(',').pop() : String(value);
}

function parseGeminiJson(rawText) {
  const trimmed = String(rawText).trim()
    .replace(/^```json\s*/i, '')
    .replace(/^```\s*/i, '')
    .replace(/```$/i, '')
    .trim();
  try {
    return JSON.parse(trimmed);
  } catch (e) {
    throw new Error('Gemini APIのJSON解析に失敗しました。応答: ' + trimmed.slice(0, 500));
  }
}

function normalizeGeminiFields(fields) {
  // 画面側が必ず candidates 配列として扱えるよう、AI応答の揺れを吸収します。
  return {
    userName: normalizeCandidates(fields.userName),
    date: normalizeCandidates(fields.date),
    summary: normalizeCandidates(fields.summary),
    amount: normalizeCandidates(fields.amount)
  };
}

function normalizeCandidates(field) {
  if (!field) return { candidates: [] };
  if (Array.isArray(field)) return { candidates: field.map(String).filter(Boolean) };
  if (field.candidates && Array.isArray(field.candidates)) {
    return { candidates: field.candidates.map(String).filter(Boolean) };
  }
  return { candidates: [String(field)].filter(Boolean) };
}

function getFileExtension(fileName, mimeType) {
  const match = String(fileName || '').match(/\.[^.]+$/);
  if (match) return match[0].toLowerCase();
  if (mimeType === 'application/pdf') return '.pdf';
  if (mimeType === 'image/jpeg') return '.jpg';
  if (mimeType === 'image/png') return '.png';
  return '';
}

function ensureReceiptsSheet(ss) {
  const sheet = ss.getSheetByName(RECEIPTS_SHEET_NAME) || ss.getSheets()[0] || ss.insertSheet(RECEIPTS_SHEET_NAME);
  if (sheet.getName() !== RECEIPTS_SHEET_NAME) sheet.setName(RECEIPTS_SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(RECEIPTS_HEADERS);
  }
  return sheet;
}
