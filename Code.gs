function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


/**
 * บันทึกข้อมูลจากฟอร์ม
 */
function saveData(obj) {
  Logger.log(obj);

  const FOLDER_ID = '1edVIlVMkNAEUn0wc7koYNU1EWx382XpZ';
  const folder = DriveApp.getFolderById(FOLDER_ID);

  let fileUrl = '';

  // ตรวจสอบว่าผู้ใช้แนบไฟล์หรือไม่
  if (obj.myFile && obj.myFile.name) {
    const file = folder.createFile(obj.myFile);
    fileUrl = file.getUrl();
  }

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheets()[0];

  sheet.appendRow([
    new Date(),
    obj.title || '',
    obj.fname || '',
    obj.lname || '',
    obj.cd || '',
    fileUrl,
    obj.wishMessage || ''
  ]);

  const evidenceText = fileUrl || 'ไม่มีไฟล์แนบ';
  const additionalText = obj.wishMessage || 'ไม่มีข้อมูลเพิ่มเติม';

  const message =
    '📢 มีผู้แจ้งข้อมูลการลา\n\n' +
    'ชื่อ: ' +
    (obj.title || '') + ' ' +
    (obj.fname || '') + ' ' +
    (obj.lname || '') + '\n' +
    'ขอลาหยุดถึงวันที่: ' + (obj.cd || '-') + '\n' +
    'ข้อมูลเพิ่มเติม: ' + additionalText + '\n' +
    'ไฟล์หลักฐาน: ' + evidenceText;

  let lineSent = false;
  let lineError = '';

  try {
    sendLineMessage(message);
    lineSent = true;
  } catch (error) {
    // ข้อมูลยังถูกบันทึกในชีต ถึงแม้ LINE ส่งไม่สำเร็จ
    lineError = error.message;
    console.error('LINE Messaging API Error: ' + error.message);
  }

  return {
    success: true,
    message: 'บันทึกข้อมูลเรียบร้อยแล้ว',
    fileUrl: fileUrl,
    lineSent: lineSent,
    lineError: lineError
  };
}


/**
 * ส่งข้อความด้วย LINE Messaging API
 */
function sendLineMessage(message) {
  const properties = PropertiesService.getScriptProperties();

  const channelAccessToken =
    properties.getProperty('LINE_CHANNEL_ACCESS_TOKEN');

  const targetId =
    properties.getProperty('LINE_TARGET_ID');

  if (!channelAccessToken) {
    throw new Error(
      'ไม่พบ LINE_CHANNEL_ACCESS_TOKEN ใน Script Properties'
    );
  }

  if (!targetId) {
    throw new Error(
      'ไม่พบ LINE_TARGET_ID ใน Script Properties'
    );
  }

  const url = 'https://api.line.me/v2/bot/message/push';

  const payload = {
    to: targetId,
    messages: [
      {
        type: 'text',
        text: message
      }
    ]
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + channelAccessToken
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const statusCode = response.getResponseCode();
  const responseText = response.getContentText();

  Logger.log('LINE status: ' + statusCode);
  Logger.log('LINE response: ' + responseText);

  if (statusCode < 200 || statusCode >= 300) {
    throw new Error(
      'ส่งข้อความ LINE ไม่สำเร็จ (' +
      statusCode +
      '): ' +
      responseText
    );
  }

  return true;
}


/**
 * ใช้ทดสอบการส่งข้อความ
 */
function testSendLineMessage() {
  sendLineMessage(
    '✅ ทดสอบส่งข้อความจาก Google Apps Script สำเร็จ'
  );
}
