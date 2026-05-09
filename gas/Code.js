/**
 * みんなの菜園会計自動化 GAS スクリプト
 * 目的：レシート画像を Gemini で解析し、「明細」シートに支出として登録
 *       正常時は 02_支出保存用画像、エラー時は 03_取込エラー画像 へ移動
 * 実行頻度：1時間に1回（時間トリガーで実行）
 */

const SETTINGS_SHEET_NAME = '設定';
const DETAILS_SHEET_NAME = '明細';
const MODEL_NAME = 'gemini-2.5-flash';
const GEMINI_ENDPOINT = 'https://generativelanguage.googleapis.com/v1beta/models/' + MODEL_NAME + ':generateContent';

function processReceiptImages() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const settings = loadSettings(spreadsheet);
  const inputFolder = getFolderById(settings.inputFolderId);
  const savedFolder = getFolderById(settings.savedFolderId);
  const errorFolder = getFolderById(settings.errorFolderId);
  const categories = loadCategoryList(spreadsheet);

  const files = inputFolder.getFiles();
  const errors = [];

  while (files.hasNext()) {
    const file = files.next();
    const originalName = file.getName();

    if (isDuplicateFile(originalName, savedFolder, errorFolder)) {
      moveToError(file, errorFolder, 'ERROR_DUPLICATE_' + originalName);
      errors.push('重複ファイル検出: ' + originalName);
      continue;
    }

    try {
      const rawText = analyzeReceiptWithGemini(file, settings.apiKey, categories);
      const parsed = parseReceiptJson(rawText);
      const rowNumber = writeExpenseRow(spreadsheet, parsed, file.getUrl());
      const newName = buildSavedFileName(rowNumber, parsed, file.getName());
      moveToFolder(file, savedFolder, newName);
    } catch (e) {
      const errorName = 'ERROR_' + originalName;
      moveToError(file, errorFolder, errorName);
      errors.push('処理失敗: ' + originalName + ' → ' + e.message);
    }
  }

  if (errors.length > 0) {
    sendErrorNotification(errors);
  }
}

function loadSettings(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(SETTINGS_SHEET_NAME);
  if (!sheet) throw new Error('設定シートが見つかりません: ' + SETTINGS_SHEET_NAME);

  const values = sheet.getDataRange().getValues();
  const map = {};

  values.forEach(function (row) {
    const key = row[0];
    const value = row[1];
    if (key && value !== undefined && value !== null && value !== '') {
      map[key.toString().trim()] = value.toString().trim();
    }
  });

  return {
    apiKey: map['Gemini APIキー'] || map['Gemini APIキー '],
    inputFolderId: map['01_支出記入用画像フォルダID'],
    savedFolderId: map['02_支出保存用画像フォルダID'],
    errorFolderId: map['03_取込エラー画像フォルダID'],
    spreadsheetId: map['スプレッドシートID'] || spreadsheet.getId(),
  };
}

function loadCategoryList(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(SETTINGS_SHEET_NAME);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  const categories = [];

  for (let row = 2; row <= lastRow; row++) {
    const value = sheet.getRange(row, 6).getValue();
    if (!value) continue;
    categories.push(value.toString().trim());
  }

  return categories.filter(function (category) {
    return category !== '';
  });
}

function getFolderById(folderId) {
  try {
    return DriveApp.getFolderById(folderId);
  } catch (e) {
    throw new Error('フォルダIDが無効です: ' + folderId);
  }
}

function isDuplicateFile(fileName, savedFolder, errorFolder) {
  return savedFolder.getFilesByName(fileName).hasNext() || errorFolder.getFilesByName(fileName).hasNext();
}

function analyzeReceiptWithGemini(file, apiKey, categories) {
  const imageBase64 = Utilities.base64Encode(file.getBlob().getBytes());
  const prompt = buildGeminiPrompt(categories);

  const requestBody = {
    contents: [{
      parts: [
        { text: prompt },
        { inline_data: { mime_type: file.getMimeType(), data: imageBase64 } },
      ],
    }],
    generationConfig: {
      temperature: 0.0,
      maxOutputTokens: 2048,
    },
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(requestBody),
    muteHttpExceptions: true,
  };

  const response = UrlFetchApp.fetch(GEMINI_ENDPOINT + '?key=' + encodeURIComponent(apiKey), options);
  const status = response.getResponseCode();
  const text = response.getContentText();

  if (status !== 200) {
    throw new Error('Gemini APIエラー: ' + status + ' / ' + text);
  }

  const parsed = JSON.parse(text);
  const outputText = extractOutputText(parsed);
  if (!outputText) {
    throw new Error('Geminiレスポンスからテキストを抽出できませんでした');
  }

  return outputText;
}

function buildGeminiPrompt(categories) {
  const categoryList = categories.length > 0 ? categories.join(' | ') : '種苗代 | 土・資材 | 農具・備品 | イベント費 | 消耗品 | 会費 | イベント収入 | 補助金・助成金 | その他';
  return [
    'あなたはGoogle Geminiです。以下のルールに従って、添付されたレシート画像の情報をJSON形式で出力してください。',
    '1. 出力は必ずJSONオブジェクトのみとし、余計な説明を含めないこと。',
    '2. フィールドは必ず次のキーを含めること: date, amount, category, purpose, note',
    '3. date は 2026-03-15 のような ISO 形式、なければ空文字列。',
    '4. amount は税込金額の整数のみ。通貨記号や小数点は除く。見つからなければ空文字列。',
    '5. category は以下のリストから最適なものを1つだけ選ぶこと: ' + categoryList + '。該当がなければ "その他"。',
    '6. purpose は用途・内容を日本語で簡潔に記載する（以下のルールに従う）：',
    '   - 同じ種類の品目はまとめて表記（例：「野菜苗、ビニール袋」）',
    '   - 日用品（袋・タオル・カップ等）は「消耗品」等の総称で表記してもよい',
    '   - 高額なものや他と種類が異なる特殊なものは個別に記載（例：「野菜苗、シャベル」）',
    '   - 全体で3〜4点程度にまとめる',
    '7. note は必要なら補足情報を入れる。',
    '8. 空欄の項目は必ず空文字列として出力する。',
    '画像の情報から日付、金額、用途を正確に抽出してください。',
  ].join('\n');
}

function extractOutputText(responseObject) {
  if (!responseObject || !responseObject.candidates || !responseObject.candidates.length) {
    return null;
  }
  const candidate = responseObject.candidates[0];
  if (candidate.content && candidate.content.parts && candidate.content.parts.length) {
    const firstPart = candidate.content.parts[0];
    if (firstPart && firstPart.text) {
      return firstPart.text;
    }
  }
  if (candidate.output) {
    return candidate.output;
  }
  if (candidate.text) {
    return candidate.text;
  }
  return null;
}

function parseReceiptJson(rawText) {
  const jsonText = extractJson(rawText.trim());
  if (!jsonText) {
    throw new Error('JSON抽出に失敗しました: ' + rawText);
  }

  let parsed;
  try {
    parsed = JSON.parse(jsonText);
  } catch (e) {
    throw new Error('JSON解析エラー: ' + e.message + '\n' + jsonText);
  }

  return {
    date: normalizeDate(parsed.date || ''),
    amount: normalizeAmount(parsed.amount || ''),
    category: parsed.category ? parsed.category.toString().trim() : '',
    purpose: parsed.purpose ? parsed.purpose.toString().trim() : '',
    note: parsed.note ? parsed.note.toString().trim() : '',
  };
}

function extractJson(text) {
  const fencedMatch = text.match(/```json\s*([\s\S]*?)\s*```/i);
  if (fencedMatch && fencedMatch[1]) {
    return fencedMatch[1].trim();
  }
  const match = text.match(/\{[\s\S]*\}/m);
  return match ? match[0] : null;
}

function normalizeDate(dateText) {
  if (!dateText) return '';
  const text = dateText.toString().trim();
  const iso = text.replace(/[年月\.\/]/g, '-').replace(/日/g, '');
  const match = iso.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (!match) return '';
  const year = match[1];
  const month = ('0' + match[2]).slice(-2);
  const day = ('0' + match[3]).slice(-2);
  return year + '-' + month + '-' + day;
}

function normalizeAmount(amountText) {
  if (amountText === undefined || amountText === null) return '';
  const digits = amountText.toString().replace(/[^0-9]/g, '');
  return digits === '' ? '' : digits;
}

function writeExpenseRow(spreadsheet, data, fileUrl) {
  const sheet = spreadsheet.getSheetByName(DETAILS_SHEET_NAME);
  if (!sheet) throw new Error('明細シートが見つかりません: ' + DETAILS_SHEET_NAME);

  const lastRow = sheet.getLastRow();
  const headerRows = 3;
  const dataRowCount = Math.max(0, lastRow - headerRows);
  const nextNo = dataRowCount + 1;
  const targetRow = lastRow + 1;
  const values = [
    nextNo,
    data.date || '',
    '支出',
    data.category || '',
    data.purpose || '',
    data.amount || '',
    fileUrl || '',
  ];

  sheet.getRange(targetRow, 1, 1, values.length).setValues([values]);
  return nextNo;
}

function buildSavedFileName(rowNumber, data, originalName) {
  const extension = originalName.lastIndexOf('.') > -1 ? originalName.slice(originalName.lastIndexOf('.')) : '';
  const datePart = data.date ? data.date.replace(/-/g, '') : 'unknown';
  const categoryPart = data.category ? sanitizeForFilename(data.category) : '不明';
  const amountPart = data.amount ? data.amount : 'unknown';
  return 'No' + rowNumber + '_' + datePart + '_' + categoryPart + '_' + amountPart + '円' + extension;
}

function sanitizeForFilename(text) {
  return text.replace(/[^0-9a-zA-Z\u4e00-\u9faf\u3040-\u309f\u30a0-\u30ff_\-]/g, '_');
}

function moveToFolder(file, folder, newName) {
  file.moveTo(folder);
  file.setName(newName);
}

function moveToError(file, errorFolder, newName) {
  moveToFolder(file, errorFolder, newName);
}

function sendErrorNotification(errors) {
  const subject = 'みんなの菜園会計：画像処理エラー通知';
  const body = '以下のエラーが発生しました。\n\n' + errors.join('\n');
  GmailApp.sendEmail(Session.getActiveUser().getEmail(), subject, body);
}

function runHourly() {
  processReceiptImages();
}
