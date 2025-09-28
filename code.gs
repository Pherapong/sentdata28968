// ============== การตั้งค่า ==============
const SHEET_ID = '10xNnXNyUjOUmuRLrQ8k08pRRIjjuR7SxV-4TRt3r3Z8'; // Sheet ID หลัก
const FOLDER_NAME = 'โฟลเดอร์ไม่มีชื่อ1'; // ชื่อ folder สำหรับเก็บไฟล์
// =====================================

function doGet(e) {
  let page = e.parameter.page;
  let docNumber = e.parameter.doc || '';
  
  let template;
  if (page === 'upload') {
    template = HtmlService.createTemplateFromFile('upload');
    template.docNumber = docNumber;
  } else {
    // หน้าเริ่มต้นคือ Index (Dashboard + Form)
    template = HtmlService.createTemplateFromFile('index');
  }
  
  template.webAppUrl = ScriptApp.getService().getUrl();
  return template.evaluate().setTitle('ระบบออกเลขหนังสือ');
}

/**
 * บันทึกข้อมูลและสร้างเลขหนังสือใหม่
 * @param {object} data ข้อมูลจากฟอร์ม
 * @returns {object} ผลลัพธ์การบันทึก
 */
function saveData(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const bookSheet = ss.getSheetByName('book_id');
  const sentSheet = ss.getSheetByName('book_sent');
  const lockSheet = ss.getSheetByName('lock') || ss.insertSheet('lock');

  if (!bookSheet) throw new Error('Sheet book_id not found');
  if (!sentSheet) throw new Error('Sheet book_sent not found');

  // ระบบ locking เพื่อป้องกัน race condition
  const maxRetries = 5;
  const lockTimeout = 10000; // 10 วินาที
  let lockAcquired = false;
  let retries = 0;

  while (!lockAcquired && retries < maxRetries) {
    try {
      // ตรวจสอบ lock ที่มีอยู่
      const currentTime = new Date().getTime();
      const lockData = lockSheet.getRange(1, 1, 1, 2).getValues()[0];
      const existingLock = lockData[0];
      const lockTime = lockData[1] ? new Date(lockData[1]).getTime() : 0;

      // ถ้ามี lock และยังไม่หมดอายุ ให้รอ
      if (existingLock && (currentTime - lockTime < lockTimeout)) {
        retries++;
        if (retries >= maxRetries) {
          return { success: false, message: 'ระบบกำลังทำงาน กรุณาลองใหม่อีกครั้งในอีกสักครู่' };
        }
        Utilities.sleep(200 + Math.random() * 300); // รอแบบสุ่ม 200-500ms
        continue;
      }

      // พยายามได้ lock
      const lockId = Utilities.getUuid();
      lockSheet.getRange(1, 1).setValue(lockId);
      lockSheet.getRange(1, 2).setValue(new Date());
      
      // รอสักนิดแล้วตรวจสอบว่าได้ lock จริงหรือไม่
      Utilities.sleep(100);
      const verifyLock = lockSheet.getRange(1, 1).getValue();
      
      if (verifyLock === lockId) {
        lockAcquired = true;
      } else {
        retries++;
        if (retries >= maxRetries) {
          return { success: false, message: 'ระบบกำลังทำงาน กรุณาลองใหม่อีกครั้ง' };
        }
        Utilities.sleep(200 + Math.random() * 300);
        continue;
      }

    } catch (error) {
      retries++;
      if (retries >= maxRetries) {
        return { success: false, message: 'เกิดข้อผิดพลาด กรุณาลองใหม่อีกครั้ง' };
      }
      Utilities.sleep(200 + Math.random() * 300);
      continue;
    }
  }

  if (!lockAcquired) {
    return { success: false, message: 'ไม่สามารถประมวลผลได้ในขณะนี้ กรุณาลองใหม่อีกครั้ง' };
  }

  try {
    const timestamp = new Date();

    // Format timestamp in Thai date and time
    const thaiMonths = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];
    const thaiYear = timestamp.getFullYear() + 543;
    const thaiDate = `${timestamp.getDate()} ${thaiMonths[timestamp.getMonth()]} ${thaiYear}`;
    const thaiTime = `${timestamp.getHours().toString().padStart(2, '0')}.${timestamp.getMinutes().toString().padStart(2, '0')}`;
    const thaiDateTime = `${thaiDate} ${thaiTime}`;

    // Format the provided date (data.date) to Thai date format
    const inputDate = new Date(data.date);
    const inputThaiDate = `${inputDate.getDate()} ${thaiMonths[inputDate.getMonth()]} ${inputDate.getFullYear() + 543}`;

    // Handle file attachments
    let fileLinks = '';
    if (data.attachments && data.attachments.length > 0) {
      const folder = getOrCreateFolder();
      const fileUrls = [];
      
      for (let i = 0; i < data.attachments.length; i++) {
        const attachment = data.attachments[i];
        try {
          // แปลง base64 เป็น Blob
          const blob = Utilities.newBlob(
            Utilities.base64Decode(attachment.data.split(',')[1]),
            attachment.type,
            attachment.name
          );
          
          // สร้างไฟล์ใน Google Drive
          const file = folder.createFile(blob);
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          
          const fileUrl = file.getUrl();
          fileUrls.push(`${attachment.name}: ${fileUrl}`);
        } catch (fileError) {
          console.error('Error uploading file:', fileError);
          fileUrls.push(`${attachment.name}: อัพโหลดไม่สำเร็จ`);
        }
      }
      fileLinks = fileUrls.join('\n');
    }

    // บันทึกข้อมูลใน book_id (เพิ่ม column สำหรับ attachments)
    bookSheet.appendRow([timestamp, data.date, data.from, data.to, data.subject, data.action || data.operator || '', fileLinks]);

    // Handle book_sent sheet - รับเลขถัดไป
    const lastRow = sentSheet.getLastRow();
    let lastNumber = 0;
    if (lastRow > 0) {
      const lastValue = sentSheet.getRange(lastRow, 1).getValue();
      lastNumber = parseInt(lastValue, 10) || 0;
    }

    const newNumber = (lastNumber + 1).toString().padStart(4, '0');
    sentSheet.appendRow([newNumber, thaiDateTime, inputThaiDate, data.from, data.to, data.subject, data.action || data.operator || '', fileLinks]);

    return { success: true, number: newNumber };

  } catch (error) {
    console.error('Error in saveData:', error);
    return { success: false, message: 'เกิดข้อผิดพลาดในการบันทึกข้อมูล กรุณาลองใหม่อีกครั้ง' };
  } finally {
    // ปล่อย lock เสมอ
    try {
      lockSheet.getRange(1, 1, 1, 2).clearContent();
    } catch (e) {
      // ถ้าไม่สามารถปล่อย lock ได้ ก็ไม่เป็นไร เพราะมี timeout อยู่แล้ว
    }
  }
}

/**
 * สร้างและบันทึกเลขหนังสือใหม่ (รองรับทั้ง 2 รูปแบบ)
 * @param {object} formData ข้อมูลที่ส่งมาจากฟอร์ม
 * @returns {string} เลขหนังสือที่สร้างขึ้นใหม่
 */
function generateAndSaveDocNumber(formData) {
  // ใช้ฟังก์ชัน saveData ที่มีอยู่แล้ว
  const result = saveData(formData);
  
  if (result.success) {
    return result.number;
  } else {
    throw new Error(result.message);
  }
}

/**
 * ดึงข้อมูลเอกสารจาก book_sent sheet
 * @returns {Array<object>} ข้อมูลเอกสารทั้งหมด
 */
function getDocuments() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('book_sent');
  
  if (!sheet || sheet.getLastRow() <= 1) return [];
  
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  
  const result = data.map(row => {
    // แปลงวันที่ไทยเป็นรูปแบบที่แสดงได้
    let docDate = '';
    if (row[2]) {
      docDate = row[2].toString();
    }
    
    // ตรวจสอบสถานะจากไฟล์แนบ
    let status = 'ยังไม่อัปโหลด';
    let fileUrl = '';
    if (row[7] && row[7].toString().trim() !== '') {
      status = 'อัปโหลดแล้ว';
      // ดึง URL แรกจากไฟล์แนบ
      const fileLinks = row[7].toString().split('\n');
      if (fileLinks.length > 0 && fileLinks[0].includes(':')) {
        fileUrl = fileLinks[0].split(': ')[1] || '';
      }
    }
    
    return {
      docNumber: row[0] || '',
      createTime: row[1] || '',
      docDate: docDate,
      from: row[3] || '',
      to: row[4] || '',
      subject: row[5] || '',
      operator: row[6] || '',
      status: status,
      fileUrl: fileUrl
    };
  }).reverse(); // แสดงรายการล่าสุดก่อน
  
  return result;
}

/**
 * อัปโหลดไฟล์และอัปเดตข้อมูลในชีต
 * @param {object} formObject ข้อมูลไฟล์
 * @returns {object} ผลลัพธ์การอัปโหลด
 */
function uploadFile(formObject) {
  try {
    const { docNumber, fileData, mimeType, fileName } = formObject;
    
    if (!fileData || !docNumber) {
      throw new Error('ข้อมูลไฟล์หรือเลขหนังสือไม่ถูกต้อง');
    }
    
    // แปลง base64 เป็น Blob
    const decodedData = Utilities.base64Decode(fileData);
    const finalFileName = fileName || `${docNumber}.pdf`;
    const blob = Utilities.newBlob(decodedData, mimeType || 'application/pdf', finalFileName);
    
    // สร้างไฟล์ใน Drive
    const folder = getOrCreateFolder();
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const fileUrl = file.getUrl();
    
    // อัปเดตข้อมูลใน book_sent sheet
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('book_sent');
    const data = sheet.getRange("A:A").getValues();
    let rowToUpdate = -1;
    
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] == docNumber) {
        rowToUpdate = i + 1;
        break;
      }
    }
    
    if (rowToUpdate !== -1) {
      // อัปเดตคอลัมน์ที่ 8 (ไฟล์แนบ)
      const currentAttachments = sheet.getRange(rowToUpdate, 8).getValue();
      let newAttachments = `${finalFileName}: ${fileUrl}`;
      
      if (currentAttachments && currentAttachments.toString().trim() !== '') {
        newAttachments = currentAttachments + '\n' + newAttachments;
      }
      
      sheet.getRange(rowToUpdate, 8).setValue(newAttachments);
    } else {
      throw new Error('ไม่พบเลขหนังสือนี้ในระบบ');
    }
    
    return { success: true, url: fileUrl, message: 'อัปโหลดไฟล์สำเร็จ' };
  } catch (e) {
    console.error('Error in uploadFile:', e);
    return { success: false, message: e.toString() };
  }
}

// สร้างหรือหา folder สำหรับเก็บไฟล์แนบ
function getOrCreateFolder() {
  const folders = DriveApp.getFoldersByName(FOLDER_NAME);
  
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return DriveApp.createFolder(FOLDER_NAME);
  }
}

/**
 * ดึงตัวเลือกสำหรับ dropdown "ส่งถึง"
 * @returns {Array<string>} รายชื่อตัวเลือก
 */
function getSentOptions() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('data');

  if (!sheet) return [];

  try {
    const values = sheet.getRange('A2:A').getValues().flat().filter(String);
    return values;
  } catch (error) {
    console.error('Error getting sent options:', error);
    return [];
  }
}

/**
 * ดึงตัวเลือกสำหรับ dropdown "จาก"
 * @returns {Array<string>} รายชื่อตัวเลือก
 */
function getFromOptions() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('data');

  if (!sheet) return [];

  try {
    const values = sheet.getRange('B2:B').getValues().flat().filter(String);
    return values;
  } catch (error) {
    console.error('Error getting from options:', error);
    return [];
  }
}