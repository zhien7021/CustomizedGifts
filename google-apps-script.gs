/**
 * 校慶客製小卡產生器 - Google 雲端硬碟接收腳本 (Google Apps Script)
 *
 * 使用方式：
 * 1. 前往 https://script.google.com/ 建立「新專案」
 * 2. 將本檔案全部內容貼上並儲存
 * 3. 點擊右上角「部署」>「新增部署作業」
 * 4. 類型選擇「網頁應用程式」：
 *    - 執行身分：我
 *    - 誰可以存取：所有人
 * 5. 授權完成後複製「網頁應用程式網址」(Web App URL)
 * 6. 將該網址貼到 index.html / 壓克力小卡產生器.html 的 GAS_WEB_APP_URL 變數中
 */

// 選填：若想指定存到某個現有雲端資料夾，請填入網址最後一段的資料夾 ID；留空會自動建立「校慶客製小卡訂單」資料夾
const FOLDER_ID = '';
const DEFAULT_FOLDER_NAME = '校慶客製小卡訂單';
const ORDER_SHEET_NAME = '校慶客製小卡訂單紀錄';

function getTargetFolder_() {
  if (FOLDER_ID) {
    return DriveApp.getFolderById(FOLDER_ID);
  }
  const folders = DriveApp.getFoldersByName(DEFAULT_FOLDER_NAME);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(DEFAULT_FOLDER_NAME);
}

function getOrCreateOrderSheet_(folder) {
  const files = folder.getFilesByName(ORDER_SHEET_NAME);
  if (files.hasNext()) {
    return SpreadsheetApp.open(files.next()).getSheets()[0];
  }
  const ss = SpreadsheetApp.create(ORDER_SHEET_NAME);
  const sheet = ss.getSheets()[0];
  sheet.appendRow(['送出時間', '班級', '姓名', '模板款式', '卡片文字', '檔案名稱', '雲端連結']);
  sheet.setFrozenRows(1);
  // 將試算表移動到同一個訂單資料夾
  const ssFile = DriveApp.getFileById(ss.getId());
  ssFile.moveTo(folder);
  return sheet;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse_({ ok: false, error: '未收到上傳資料' });
    }

    const data = JSON.parse(e.postData.contents);
    if (!data.base64 || !data.filename) {
      return jsonResponse_({ ok: false, error: '缺少圖片內容或檔名' });
    }

    const folder = getTargetFolder_();
    const bytes = Utilities.base64Decode(data.base64);
    const blob = Utilities.newBlob(bytes, data.mimeType || 'image/png', data.filename);
    const file = folder.createFile(blob);

    lock.waitLock(15000);
    try {
      const sheet = getOrCreateOrderSheet_(folder);
      const timeStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');
      sheet.appendRow([
        timeStr,
        data.customerClass || '',
        data.customerName || '',
        data.preset || '',
        data.cardLabel || '',
        file.getName(),
        file.getUrl(),
      ]);
    } finally {
      lock.releaseLock();
    }

    return jsonResponse_({
      ok: true,
      fileName: file.getName(),
      fileUrl: file.getUrl(),
    });
  } catch (err) {
    return jsonResponse_({
      ok: false,
      error: err && err.message ? err.message : String(err),
    });
  }
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
