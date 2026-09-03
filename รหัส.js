 // =========================================================================
// 🚀 ระบบหลังบ้าน (Backend) - โครงสร้างตรงกับ Sheet ที่ออกแบบใหม่ทั้งหมด
// =========================================================================

// เชื่อมต่อกับ Sheet ID ของคุณครูโดยตรง
const SPREADSHEET_ID = "14L_gXjHCl9PDZ_ap95A8uYu3mfbm-0LvzBxSKXMqFyI"; 
const FOLDER_ID = '1MzU1ZQsHfKHAtoLSeWl_I7WyP--Wufcd';

const FOLDER_NAME = "LibrarySystem_Files";

function getSpreadsheet() {
  if (SPREADSHEET_ID !== "") {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getOrCreateSheet(sheetName, headers) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);

    if (headers && headers.length > 0) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight("bold")
        .setBackground("#e0f2fe");

      sheet.setFrozenRows(1);
    }
  }

  return sheet;
}

// 🛠️ 1. กดปุ่ม "เรียกใช้" ที่ฟังก์ชันนี้ เพื่อสร้างฐานข้อมูล
function setupSystem() {
  // โครงสร้างชีตใหม่ทั้งหมด (ไม่มี Config, Borrows, ReadingLogs, Homeworks)
  getOrCreateSheet('students', ['ชั้น', 'ชื่อนักเรียน']);
  getOrCreateSheet('data', ['วันที่บันทึก', 'ชั้น', 'ชื่อ-สกุล', 'เพศ', 'วันเกิด', 'ชื่อเล่น', 'เบอร์โทร', 'ลิงก์รูปถ่าย']);
  getOrCreateSheet('Users', ['วันที่ลงทะเบียน', 'ชั้น', 'ชื่อ-สกุล', 'FaceData']);
  getOrCreateSheet('Attendance', ['ชื่อ-สกุล', 'ชั้น', 'วันที่', 'เวลา']);
  getOrCreateSheet('borrow', ['วันที่บันทึก', 'ชั้น', 'ชื่อ-สกุล', 'ชื่อหนังสือ', 'วันที่ยืม', 'กำหนดคืน', 'สถานะ']);
  getOrCreateSheet('reading', ['วันที่ส่ง', 'ชั้น', 'ชื่อ-สกุล', 'ชื่อเรื่อง', 'เรื่องย่อ', 'คะแนน(ดาว)', 'สถานะ']);
  getOrCreateSheet('homework', ['วันที่ส่ง', 'ชั้น', 'ชื่อคนที่1', 'ชื่อคนที่2', 'ชื่องาน', 'ลิงก์ไฟล์1', 'ลิงก์ไฟล์2', 'สถานะ']);
  getOrCreateSheet('link', ['วันที่เวลา', 'ชื่อลิ้ง', 'url']);
  getOrCreateSheet('dictionary', ['คำที่เขียนผิด', 'คำที่ถูกต้อง', 'หมวดหมู่/หมายเหตุ']);
  getOrCreateSheet('Assignments', ['AssignmentID', 'Title', 'Content', 'TargetClass', 'CreatedAt']);
  getOrCreateSheet('HwTopics', ['TopicID', 'Title', 'TargetClass', 'CreatedAt']);
  getOrCreateSheet('lessons', ['LessonNo', 'Topic', 'ExternalLink', 'ContentType', 'TargetClass', 'CreatedAt']);
  var configSheet = getOrCreateSheet('Config', ['KeyName', 'KeyValue', 'Description']);
  if (configSheet.getLastRow() <= 1) {
    configSheet.appendRow(['GAS_API_URL', 'https://script.google.com/macros/s/AKfycbwg-QW1zSlqghhLTGz3EDZMDDw2nAf72uuWygYzoEJxGFEF7pnZRoAqk0WNiZfvXvxClw/exec', 'ลิงก์ระบบหลังบ้าน Web App สำหรับใช้เชื่อมต่อ']);
  }
  
  // ชีตสำหรับให้คุณครูนำไปดึงสูตรสถิติต่อ
  getOrCreateSheet('sum', ['สถิติ (รอตั้งค่าสูตร)']);
  getOrCreateSheet('number 5', ['5 อันดับรักการอ่าน (รอตั้งค่าสูตร)']);
  getOrCreateSheet('Stats_Counter', ['Counters (รอตั้งค่าสูตร)']);
  
  var folders = DriveApp.getFoldersByName(FOLDER_NAME);
  if (!folders.hasNext()) DriveApp.createFolder(FOLDER_NAME);

  populateDictionary200Words();
}

// =========================================================================
// 📡 2. ส่วนรับการดึงข้อมูล (GET Requests) 
// =========================================================================
function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : null;

    if (!action) {
      return ContentService.createTextOutput(JSON.stringify({ 
        status: "success", 
        message: "✅ เชื่อมต่อ Google Apps Script สำเร็จ 100%!" 
      })).setMimeType(ContentService.MimeType.JSON);
    }

    var data = null;
    
    if (action === 'getDashboardSummary') data = getDashboardSummary();
    else if (action === 'getConfig') data = getConfig();
    else if (action === 'getStudents') data = getStudents(e.parameter.class);
    else if (action === 'getAllStudents') data = getAllStudents();
    else if (action === 'getKnownFaces') data = getKnownFaces();
    else if (action === 'getStats') data = getStats();
    else if (action === 'getStudentData') data = getStudentData();
    else if (action === 'getBorrows') data = getBorrows();
    else if (action === 'getReadingLogs') data = getReadingLogs();
    else if (action === 'getReadingStats') data = getReadingStats();
    else if (action === 'getHomeworks') data = getHomeworks();
    else if (action === 'getLinks') data = getLinks();
    else if (action === 'getDictionaryWords') data = getDictionaryWords();
    else if (action === 'getAssignments') data = getAssignments(e.parameter.class);
    else if (action === 'getHwTopics') data = getHwTopics(e.parameter.class);
    else if (action === 'getLessons') data = getLessons();
    else if (action === 'getExams') data = getExams();
    else if (action === 'getExamScores') data = getExamScores();
    else data = { status: "success", message: "API เชื่อมต่อสมบูรณ์พร้อมใช้งาน!" };

    return ContentService.createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// =========================================================================
// 📨 3. ส่วนรับการบันทึกข้อมูล (POST Requests)
// =========================================================================
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "No data received" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var payload = JSON.parse(e.postData.contents);
    var action = payload.action;
    var data = { status: "success", message: "รับข้อมูลสำเร็จ" };

    if (action === 'logAttendance') data = logAttendance(payload) || data;
    else if (action === 'saveConfig') saveConfig(payload);
    else if (action === 'registerUser') registerUser(payload);
    else if (action === 'addStudentData') addStudentData(payload);
    else if (action === 'editStudentData') editStudentData(payload);
    else if (action === 'deleteStudentData') deleteStudentData(payload);
    else if (action === 'addBorrow') addBorrow(payload);
    else if (action === 'updateBorrowStatus') updateBorrowStatus(payload);
    else if (action === 'addReadingLog') addReadingLog(payload);
    else if (action === 'editReadingLog') editReadingLog(payload);
    else if (action === 'updateReadingStatus') updateReadingStatus(payload);
    else if (action === 'addHomework') addHomework(payload);
    else if (action === 'editHomework') editHomework(payload);
    else if (action === 'updateHomeworkStatus') updateHomeworkStatus(payload);
    else if (action === 'addLink') addLink(payload);
    else if (action === 'deleteLink') deleteLink(payload);
    else if (action === 'saveAssignment') saveAssignment(payload);
    else if (action === 'deleteAssignment') deleteAssignment(payload);
    else if (action === 'saveHwTopic') data = saveHwTopic(payload);
    else if (action === 'deleteHwTopic') data = deleteHwTopic(payload);
    else if (action === 'saveLesson') data = saveLesson(payload);
    else if (action === 'saveExamConfig') data = saveExamConfig(payload);
    else if (action === 'uploadExamPDF') data = uploadExamPDF(payload);
    else if (action === 'saveExamScore') data = saveExamScore(payload);
    else if (action === 'addStudent') {
      var sheet = getOrCreateSheet('students');
      // เพิ่มลง students คอลัมน์ A=ชั้น, B=ชื่อ
      sheet.appendRow([payload.className, payload.name]);
    }

    return ContentService.createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// =========================================================================
// 🗂️ 4. ฟังก์ชันจัดการข้อมูล (Data Handlers)
// =========================================================================

function uploadFileToDrive(base64Data, mimeType, fileName) {

  try {

    if (!base64Data) {
      return '-';
    }

    var folder = DriveApp.getFolderById(FOLDER_ID);

    var cleanData = base64Data;

    if (base64Data.indexOf(',') !== -1) {
      cleanData = base64Data.split(',')[1];
    }

    var bytes = Utilities.base64Decode(cleanData);

    var blob = Utilities.newBlob(
      bytes,
      mimeType || 'image/jpeg',
      fileName || ('photo_' + Date.now() + '.jpg')
    );

    var file = folder.createFile(blob);

    // ไม่ต้อง setSharing

    return file.getId();

  } catch(err) {

    Logger.log(err);

    return 'ERROR: ' + err;
  }
}

function saveConfig(p) {
  var sheet = getOrCreateSheet('Config', ['KeyName', 'KeyValue', 'Description']);
  var data = sheet.getDataRange().getValues();
  
  if (p.keyName || p.keyValue !== undefined || p.geminiApiKey !== undefined || p.apiKey !== undefined) {
    var keyName = p.keyName || 'GEMINI_API_KEY';
    var keyValue = (p.keyValue !== undefined) ? p.keyValue : ((p.geminiApiKey !== undefined) ? p.geminiApiKey : (p.apiKey || ''));
    var found = false;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === keyName) {
        sheet.getRange(i + 1, 2).setValue(keyValue);
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow([keyName, keyValue, 'ค่าตั้งค่าระบบ ' + keyName]);
    }
  }
  
  if (p.lat !== undefined && p.lng !== undefined) {
    var keys = [
      ['Target Latitude', p.lat],
      ['Target Longitude', p.lng],
      ['Allowed Radius (KM)', p.radius]
    ];
    keys.forEach(function(item) {
      var found = false;
      for (var i = 1; i < data.length; i++) {
        if (data[i][0].toString().trim() === item[0]) {
          sheet.getRange(i + 1, 2).setValue(item[1]);
          found = true;
          break;
        }
      }
      if (!found) {
        sheet.appendRow([item[0], item[1], 'ตั้งค่าพิกัด']);
      }
    });
  }

  return { status: "success", message: "บันทึกการตั้งค่าเรียบร้อยแล้ว" };
}

function testDrive() {
  DriveApp.getRootFolder();
}

function getConfig() {
  var sheet = getOrCreateSheet('Config', ['KeyName', 'KeyValue', 'Description']);
  var config = { lat: 0, lng: 0, radius: 0.5, GEMINI_API_KEY: '' };
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) {
      config[data[i][0].toString().trim()] = data[i][1] ? data[i][1].toString().trim() : '';
    }
  }
  if (data.length >= 4 && data[1][0] === 'Target Latitude') {
    config.lat = parseFloat(data[1][1]) || 0;
    config.lng = parseFloat(data[2][1]) || 0;
    config.radius = parseFloat(data[3][1]) || 0;
  }
  return config;
}

function getDashboardSummary() {
  return {
    attendance: Math.max(0, getOrCreateSheet('Attendance').getLastRow() - 1),
    borrows: Math.max(0, getOrCreateSheet('borrow').getLastRow() - 1),
    reading: Math.max(0, getOrCreateSheet('reading').getLastRow() - 1),
    homework: Math.max(0, getOrCreateSheet('homework').getLastRow() - 1)
  };
}

// --- Attendance & Stats ---
function logAttendance(p) {
  var sheet = getOrCreateSheet('Attendance');
  var d = new Date();
  var dateStr = Utilities.formatDate(d, "Asia/Bangkok", "dd/MM/yyyy");
  var timeStr = Utilities.formatDate(d, "Asia/Bangkok", "HH:mm:ss");
  
  // ป้องกันการลงชื่อซ้ำภายใน 60 นาที ในระดับ Server-side
  try {
    var data = sheet.getDataRange().getValues();
    if (data.length > 1) {
      var nowMs = d.getTime();
      var oneHourMs = 60 * 60 * 1000;
      for (var i = data.length - 1; i >= 1; i--) {
        var rowName = data[i][0];
        var rowDate = data[i][2];
        var rowTime = data[i][3];
        if (rowName && p.name && String(rowName).trim() === String(p.name).trim()) {
          var formattedDate = rowDate instanceof Date ? Utilities.formatDate(rowDate, "Asia/Bangkok", "dd/MM/yyyy") : String(rowDate);
          if (formattedDate === dateStr) {
            var formattedTime = rowTime instanceof Date ? Utilities.formatDate(rowTime, "Asia/Bangkok", "HH:mm:ss") : String(rowTime);
            var rowTimestamp = parseAttendanceTime(formattedDate, formattedTime);
            if (rowTimestamp && !isNaN(rowTimestamp.getTime()) && (nowMs - rowTimestamp.getTime() < oneHourMs)) {
              Logger.log("Duplicate attendance skipped for " + p.name);
              return { status: "success", message: "มีการลงชื่อซ้ำภายใน 1 ชั่วโมงแล้ว (ข้ามการบันทึก)" };
            }
          }
        }
      }
    }
  } catch (err) {
    Logger.log("Error checking duplicate in logAttendance: " + err);
  }

  // Attendance: A=ชื่อ, B=ชั้น, C=วันที่, D=เวลา
  sheet.appendRow([p.name, p.className, dateStr, timeStr]);
  return { status: "success", message: "บันทึกสำเร็จ" };
}

function parseAttendanceTime(dStr, tStr) {
  try {
    var dp = dStr.split('/');
    var tp = tStr.split(':');
    if (dp.length === 3 && tp.length >= 2) {
      return new Date(
        parseInt(dp[2], 10),
        parseInt(dp[1], 10) - 1,
        parseInt(dp[0], 10),
        parseInt(tp[0], 10),
        parseInt(tp[1], 10),
        tp[2] ? parseInt(tp[2], 10) : 0
      );
    }
  } catch (e) {}
  return null;
}

function getStats() {
  var data = getOrCreateSheet('Attendance').getDataRange().getValues();
  if (data.length <= 1) return { overall: [], byClass: {} };
  var counts = {};
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) {
      var key = data[i][1] + "|" + data[i][0]; // ชั้น|ชื่อ-สกุล
      counts[key] = (counts[key] || 0) + 1;
    }
  }
  var resultList = [];
  for (var k in counts) {
    var parts = k.split("|");
    resultList.push({ cls: parts[0], name: parts[1], count: counts[k] });
  }
  resultList.sort(function(a, b) { return b.count - a.count; });
  var byClass = { "ป.1": [], "ป.2": [], "ป.3": [], "ป.4": [], "ป.5": [], "ป.6": [] };
  resultList.forEach(function(item) { if (byClass[item.cls]) byClass[item.cls].push(item); });
  return { overall: resultList.slice(0, 10), byClass: byClass };
}

// --- Users (ลงทะเบียน FaceData) ---
function getKnownFaces() {
  var data = getOrCreateSheet('Users').getDataRange().getValues();
  var faces = [];
  for (var i = 1; i < data.length; i++) {
    // Users: B=ชั้น(1), C=ชื่อ(2), D=FaceData(3)
    // FaceData รองรับ 2 formats:
    //   - เก่า: [number, number, ...] (descriptor เดี่ยว)
    //   - ใหม่: [[...d1...],[...d2...],[...d3...]] (3 มุม)
    var faceStr = data[i][3];
    if (faceStr && faceStr !== '-') {
      try {
        var parsed = JSON.parse(faceStr);
        var label = data[i][2], cls = data[i][1];
        // ตรวจว่าเป็น array ของ arrays (multi-descriptor) หรือเปล่า
        if (Array.isArray(parsed) && parsed.length > 0 && Array.isArray(parsed[0])) {
          // format ใหม่: [[d1], [d2], [d3]] → แตกออกเป็นหลาย entry
          parsed.forEach(function(desc) {
            faces.push({ label: label, class: cls, descriptor: desc });
          });
        } else {
          // format เดิม: [number, ...] → ใช้ตามปกติ
          faces.push({ label: label, class: cls, descriptor: parsed });
        }
      } catch(e) {}
    }
  }
  return faces;
}


function registerUser(p) {
  var sheet = getOrCreateSheet('Users');
  var data = sheet.getDataRange().getValues();
  var newDesc = p.faceDescriptor; // array ของตัวเลข [n1, n2, ...]
  var dtStr = new Date().toISOString();

  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === p.className && data[i][2] === p.name) {
      // พบแถวนักเรียนแล้ว: อ่าน descriptors เดิม แล้วสะสมเพิ่ม (สูงสุด 3 มุม)
      var existing = [];
      try {
        var parsed = JSON.parse(data[i][3]);
        if (Array.isArray(parsed) && parsed.length > 0 && Array.isArray(parsed[0])) {
          existing = parsed; // format ใหม่: [[d1],[d2],[d3]]
        } else if (Array.isArray(parsed)) {
          existing = [parsed]; // format เก่า: [n1,n2,...] → ห่อเป็น array
        }
      } catch(e) {}
      // สะสม descriptor ใหม่ต่อท้าย (ไม่เกิน 3 ตัว)
      existing.push(newDesc);
      if (existing.length > 3) existing = existing.slice(existing.length - 3);
      sheet.getRange(i + 1, 4).setValue(JSON.stringify(existing));
      return;
    }
  }
  // ไม่พบแถวเดิม: สร้างแถวใหม่ด้วย descriptor มุมแรก
  sheet.appendRow([dtStr, p.className, p.name, JSON.stringify([newDesc])]);
}


// --- Students (รายชื่อแสดงใน Dropdown) ---
function getStudents(className) {
  var data = getOrCreateSheet('students').getDataRange().getValues();
  var names = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === className && data[i][1]) {
      names.push(data[i][1]); // A=ชั้น(0), B=ชื่อ(1)
    }
  }
  return names;
}

// ดึงรายชื่อนักเรียนทุกชั้นในครั้งเดียว (เพื่อ cache ฝั่ง Frontend)
function getAllStudents() {
  var data = getOrCreateSheet('students').getDataRange().getValues();
  var result = {};
  for (var i = 1; i < data.length; i++) {
    var cls = data[i][0], name = data[i][1];
    if (cls && name) {
      if (!result[cls]) result[cls] = [];
      result[cls].push(name);
    }
  }
  return result;
}

// --- Data (ทำเนียบนักเรียน) ---
function getStudentData() {
  var data = getOrCreateSheet('data').getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    list.push({ 
      rowIdx: i + 1, 
      className: data[i][1], // B
      name: data[i][2],      // C
      gender: data[i][3],    // D
      dob: data[i][4],       // E
      nickname: data[i][5],  // F
      phone: data[i][6],     // G
      photoId: data[i][7]    // H
    });
  }
  return list;
}

function addStudentData(p) {
  var photoId = p.photoBase64 ? uploadFileToDrive(p.photoBase64, p.photoType, p.photoName) : '-';
  var dtStr = new Date().toISOString();
  getOrCreateSheet('data').appendRow([dtStr, p.className, p.name, p.gender, p.dob, p.nickname, "'" + p.phone, photoId]);
}

function editStudentData(p) {
  var sheet = getOrCreateSheet('data'), row = parseInt(p.rowIdx), photoId = p.photoBase64 ? uploadFileToDrive(p.photoBase64, p.photoType, p.photoName) : p.oldPhotoId;
  // อัปเดตคอลัมน์ B ถึง H (เริ่มคอลัมน์ 2, จำนวน 7 คอลัมน์)
  sheet.getRange(row, 2, 1, 7).setValues([[p.className, p.name, p.gender, p.dob, p.nickname, "'" + p.phone, photoId]]);
}

function deleteStudentData(p) {
  var sheet = getOrCreateSheet('data');
  sheet.deleteRow(parseInt(p.rowIdx));
}

// --- Borrow (ยืม-คืน) ---
function getBorrows() {
  var data = getOrCreateSheet('borrow').getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][2]) list.push({ 
      rowIdx: i + 1, 
      className: data[i][1], 
      name: data[i][2], 
      bookName: data[i][3], 
      borrowDate: data[i][4], 
      returnDate: data[i][5], 
      status: data[i][6] 
    });
  }
  return list.reverse();
}

function addBorrow(p) {
  var dtStr = new Date().toISOString();
  getOrCreateSheet('borrow').appendRow([dtStr, p.className, p.name, p.bookName, p.borrowDate, p.returnDate, 'กำลังยืม']);
}

function updateBorrowStatus(p) {
  getOrCreateSheet('borrow').getRange(parseInt(p.rowIdx), 7).setValue(p.status);
}

// --- Reading Logs (รักการอ่าน) ---
function getReadingLogs() {
  var data = getOrCreateSheet('reading').getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][2]) list.push({ 
      rowIdx: i + 1, 
      className: data[i][1], 
      name: data[i][2], 
      bookTitle: data[i][3], 
      synopsis: data[i][4], 
      rating: data[i][5], 
      status: data[i][6], 
      date: data[i][0] 
    });
  }
  return list.reverse();
}

function addReadingLog(p) {
  var dtStr = new Date().toISOString();
  getOrCreateSheet('reading').appendRow([dtStr, p.className, p.name, p.bookTitle, p.synopsis, p.rating, 'รอตรวจ']);
}

function editReadingLog(p) {
  // อัปเดตคอลัมน์ B ถึง G (เริ่ม 2, จำนวน 6 คอลัมน์): className, name, bookTitle, synopsis, rating, status
  if (p.className && p.name) {
    getOrCreateSheet('reading').getRange(parseInt(p.rowIdx), 2, 1, 6).setValues([[p.className, p.name, p.bookTitle, p.synopsis, p.rating, 'รอตรวจ']]);
  } else {
    getOrCreateSheet('reading').getRange(parseInt(p.rowIdx), 4, 1, 4).setValues([[p.bookTitle, p.synopsis, p.rating, 'รอตรวจ']]);
  }
}

function updateReadingStatus(p) {
  getOrCreateSheet('reading').getRange(parseInt(p.rowIdx), 7).setValue(p.status);
}

function getReadingStats() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('reading');
  if (!sheet) return {
    overall: [],
    byClass: {}
  };

  const data = sheet.getDataRange().getValues();

  const overall = {};
  const byClass = {};

  for (let i = 1; i < data.length; i++) {
    const className = data[i][1];   // ชั้น
    const student = data[i][2];     // ชื่อ
    if (!student) continue;

    // แยกรายชื่อหากเป็นงานกลุ่ม (คั่นด้วยจุลภาค)
    var studentNames = student.split(/, |,/).map(function(s) { return s.trim(); }).filter(Boolean);
    
    studentNames.forEach(function(studentName) {
      // รวมทั้งโรงเรียน
      overall[studentName] = (overall[studentName] || 0) + 1;

      // แยกตามชั้น
      if (!byClass[className]) {
        byClass[className] = {};
      }
      byClass[className][studentName] = (byClass[className][studentName] || 0) + 1;
    });
  }

  const overallTop =
    Object.entries(overall)
      .map(([name,count])=>({name,count}))
      .sort((a,b)=>b.count-a.count);

  const overallTop10 = overallTop.slice(0, 10);
  const classTop = {};

  Object.keys(byClass).forEach(cls=>{
    classTop[cls] =
      Object.entries(byClass[cls])
      .map(([name,count])=>({name,count}))
      .sort((a,b)=>b.count-a.count)
      .slice(0, 10);
  });

  // NOTE: ไม่เขียนลงชีตใน getReadingStats เพราะทำให้ช้า
  // การ sync ลงชีต number 5 / number ควรทำผ่าน trigger แยก

  return {
    overall: overallTop10,
    byClass: classTop
  };
}

// --- Homework (ส่งงาน) ---
function getHomeworks() {
  var data = getOrCreateSheet('homework').getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][4]) list.push({ 
      rowIdx: i + 1, 
      className: data[i][1], 
      name1: data[i][2], 
      name2: data[i][3], 
      taskName: data[i][4], 
      file1: data[i][5], 
      file2: data[i][6], 
      status: data[i][7], 
      date: data[i][0] 
    });
  }
  return list.reverse();
}

function addHomework(p) {
  var f1 = p.file1 ? uploadFileToDrive(p.file1.data, p.file1.type, p.file1.name) : '-';
  var f2 = p.file2 ? uploadFileToDrive(p.file2.data, p.file2.type, p.file2.name) : '-';
  var dtStr = new Date().toISOString();
  var status = p.status || 'รอตรวจ';
  getOrCreateSheet('homework').appendRow([dtStr, p.className, p.name1, p.name2 || '-', p.taskName, f1, f2, status]);
}

function updateHomeworkStatus(p) {
  var sheet = getOrCreateSheet('homework');
  sheet.getRange(parseInt(p.rowIdx), 8).setValue(p.status);
  
  if (p.score !== undefined) {
    var currentVal = sheet.getRange(parseInt(p.rowIdx), 7).getValue();
    
    // ถ้ามีการกรอกคะแนนเข้ามา ให้เซฟทับเลย
    if (p.score !== '') {
      sheet.getRange(parseInt(p.rowIdx), 7).setValue(p.score);
    } 
    // แต่ถ้าครูไม่ได้กรอกคะแนน (เว้นว่าง) และค่าเดิม "ไม่ใช่ลิงก์" ให้รีเซ็ตกลับเป็น '-'
    // (เพื่อป้องกันการไปลบลิงก์ไฟล์ที่ 2 ของนักเรียนโดยไม่ได้ตั้งใจ)
    else if (!String(currentVal).startsWith('http')) {
      sheet.getRange(parseInt(p.rowIdx), 7).setValue('-');
    }
  }
}

// --- Links (ฝากลิงก์) ---
function getLinks() {
  var data = getOrCreateSheet('link').getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) { 
    if (data[i][1]) list.push({ rowIdx: i + 1, title: data[i][1], url: data[i][2] }); 
  }
  return list.reverse();
}

function addLink(p) { 
  var dtStr = new Date().toISOString();
  getOrCreateSheet('link').appendRow([dtStr, p.title, p.url]); 
}

function deleteLink(p) { 
  var sheet = getOrCreateSheet('link');
  if (p.rowIdx && parseInt(p.rowIdx) > 1) {
    sheet.deleteRow(parseInt(p.rowIdx));
  }
}

// --- Dictionary (คลังคำมักเขียนผิด 200 คำ) ---
function getDictionaryWords() {
  var sheet = getOrCreateSheet('dictionary', ['คำที่เขียนผิด', 'คำที่ถูกต้อง', 'หมวดหมู่/หมายเหตุ']);
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    populateDictionary200Words();
    data = sheet.getDataRange().getValues();
  }
  var list = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] || data[i][1]) {
      list.push({
        wrong: data[i][0] ? data[i][0].toString().trim() : '',
        correct: data[i][1] ? data[i][1].toString().trim() : '',
        note: data[i][2] ? data[i][2].toString().trim() : ''
      });
    }
  }
  return list;
}

function populateDictionary200Words() {
  var sheet = getOrCreateSheet('dictionary', ['คำที่เขียนผิด', 'คำที่ถูกต้อง', 'หมวดหมู่/หมายเหตุ']);
  if (sheet.getLastRow() > 1) return;
  
  var words = [
    ["กระทัดรัด", "กะทัดรัด", "ไม่มี ร ควบกล้ำ"],
    ["กระเพรา", "กะเพรา", "กะ (ไม่มี ร), เพรา (มี ร)"],
    ["กระเทย", "กะเทย", "กะ (ไม่มี ร)"],
    ["กระโหลก", "กะโหลก", "กะ (ไม่มี ร)"],
    ["กระทันหัน", "กะทันหัน", "กะ (ไม่มี ร)"],
    ["กระพริบ", "กะพริบ", "กะ (ไม่มี ร), พริบ (มี ร)"],
    ["กระละมัง", "กะละมัง", "กะ (ไม่มี ร)"],
    ["กระล่อน", "กะล่อน", "กะ (ไม่มี ร)"],
    ["กระเหรี่ยง", "กะเหรี่ยง", "กะ (ไม่มี ร)"],
    ["กระโหลกกระลา", "กะโหลกกะลา", "กะ (ไม่มี ร)"],
    ["กระตัก", "กะตัก", "กะ (ไม่มี ร)"],
    ["กระทิ", "กะทิ", "กะ (ไม่มี ร)"],
    ["กระพง", "กะพง", "กะ (ไม่มี ร)"],
    ["กระโปโล", "กะโปโล", "กะ (ไม่มี ร)"],
    ["กระหนุงกระหนิง", "กะหนุงกะหนิง", "กะ (ไม่มี ร)"],
    ["กะวนกะวาย", "กระวนกระวาย", "มี ร ควบกล้ำ"],
    ["กะจุ๋มกะจิ๋ม", "กระจุ๋มกระจิ๋ม", "มี ร ควบกล้ำ"],
    ["กะจุกกะจิก", "กระจุกกระจิก", "มี ร ควบกล้ำ"],
    ["กะตือรือร้น", "กระตือรือร้น", "มี ร ควบกล้ำ"],
    ["กะถด", "กระถด", "มี ร ควบกล้ำ"],
    ["กะหน่ำ", "กระหน่ำ", "มี ร ควบกล้ำ"],
    ["สังเกตุ", "สังเกต", "ไม่มี สระอุ"],
    ["อนุญาติ", "อนุญาต", "ไม่มี สระอิ"],
    ["โควต้า", "โควตา", "ไม่มีไม้โท"],
    ["ซีรี่ย์", "ซีรีส์", "ส เสือ เป็นตัวสะกด"],
    ["ซีรี่ส์", "ซีรีส์", "ส เสือ เป็นตัวสะกด"],
    ["คลินิค", "คลินิก", "ก ไก่ เป็นตัวสะกด"],
    ["กราฟฟิก", "กราฟิก", "ฟ ฟัน ตัวเดียว"],
    ["ลิงค์", "ลิงก์", "ก ไก่ การันต์"],
    ["ลิ้งค์", "ลิงก์", "ก ไก่ การันต์"],
    ["อลุ่มอล่วย", "อะลุ้มอล่วย", "อะ-ลุ้ม-อะ-ล่วย"],
    ["ผัดไท", "ผัดไทย", "มี ย ยักษ์"],
    ["คำนวน", "คำนวณ", "ใช้ ณ สะกด"],
    ["คำนวณการ", "คำนวณการ", "ใช้ ณ สะกด"],
    ["เซ็นต์ชื่อ", "เซ็นชื่อ", "ไม่มี ต การันต์"],
    ["เปอร์เซนต์", "เปอร์เซ็นต์", "มี ไต่คู้ และ ต การันต์"],
    ["โน๊ต", "โน้ต", "ใช้ไม้โท"],
    ["เค็ก", "เค้ก", "ใช้ไม้โท"],
    ["เสื้อเชิ๊ต", "เสื้อเชิ้ต", "ใช้ไม้โท"],
    ["แก๊งค์", "แก๊ง", "ไม่มี ค การันต์"],
    ["เกมส์", "เกม", "ไม่มี ส การันต์"],
    ["เว็ปไซด์", "เว็บไซต์", "ส เสือ และ ต การันต์"],
    ["เว็บไซท์", "เว็บไซต์", "ใช้ ต การันต์"],
    ["อินเตอร์เน็ต", "อินเทอร์เน็ต", "ท ทหาร และ ต การันต์"],
    ["อัพเดท", "อัปเดต", "ป ปลา และ ต การันต์"],
    ["ดิจิตอล", "ดิจิทัล", "ท ธง และ ล ลิง"],
    ["อีเมล์", "อีเมล", "ไม่มี ล การันต์"],
    ["สแกนน์", "สแกน", "ไม่มี น การันต์"],
    ["แอพพลิเคชั่น", "แอปพลิเคชัน", "ป ปลา และ น หนู"],
    ["สมารทโฟน", "สมาร์ทโฟน", "มี ร การันต์"],
    ["โปรเจค", "โพรเจกต์", "หรือ โปรเจกต์"],
    ["โปรเจ็ค", "โพรเจกต์", "หรือ โปรเจกต์"],
    ["โปรโมชั่น", "โปรโมชัน", "น หนู เป็นตัวสะกด"],
    ["ฟังก์ชั่น", "ฟังก์ชัน", "น หนู เป็นตัวสะกด"],
    ["คอนเซ็ปต์", "คอนเซปต์", "ไม่มีไม้ไต่คู้"],
    ["คอนเซ็ป", "คอนเซปต์", "มี ต การันต์"],
    ["แฟชชั่น", "แฟชั่น", "ช ช้าง ตัวเดียว"],
    ["คอรัปชั่น", "คอร์รัปชัน", "มี ร การันต์"],
    ["สเป็ค", "สเปก", "ไม่มีไม้ไต่คู้ และ ก ไก่"],
    ["เช็ค", "เช็ก", "ก ไก่ สะกด"],
    ["บล๊อก", "บล็อก", "ใช้ไม้ไต่คู้"],
    ["ล็อค", "ล็อก", "ก ไก่ สะกด"],
    ["คลิ๊ก", "คลิก", "ไม่มีไม้ไต่คู้"],
    ["พล๊อต", "พล็อต", "ใช้ไม้ไต่คู้"],
    ["แร็กเก็ต", "แร็กเกต", "ไม่มีไม้ไต่คู้"],
    ["แทบเล็ต", "แท็บเล็ต", "มีไม้ไต่คู้"],
    ["ช้อปปิ้ง", "ชอปปิง", "ไม่มีไม้โท"],
    ["แซนวิช", "แซนด์วิช", "มี ด การันต์"],
    ["สเต็ก", "สเต๊ก", "ใช้ไม้ไต่คู้"],
    ["บุฟเฟ่ต์", "บุฟเฟต์", "ไม่มีไม้เอก"],
    ["เสริฟ", "เสิร์ฟ", "มี ร การันต์"],
    ["บาร์บีคิว", "บาร์เบคิว", "คำยืมทับศัพท์"],
    ["แอลกอฮอร์", "แอลกอฮอล์", "ใช้ ล การันต์"],
    ["พลาสติค", "พลาสติก", "ก ไก่ สะกด"],
    ["แท๊กซี่", "แท็กซี่", "ใช้ไม้ไต่คู้"],
    ["ไซด์", "ไซต์", "ต การันต์ (สถานที่)"],
    ["ออฟฟิต", "ออฟฟิศ", "ศ ศาลา สะกด"],
    ["ไนท์คลับ", "ไนต์คลับ", "ต ตระกร้า/ต เต่า"],
    ["รถเมล", "รถเมล์", "มี ล การันต์"],
    ["มอเตอร์ไซ", "มอเตอร์ไซค์", "มี ค การันต์"],
    ["ก็อกน้ำ", "ก๊อกน้ำ", "ใช้ไม้ตรี"],
    ["บัลลัง", "บัลลังก์", "มี ก การันต์"],
    ["ประสพภัย", "ประสบภัย", "บ ใบไม้ สะกด"],
    ["มนุษย์สัมพันธ์", "มนุษยสัมพันธ์", "ไม่มี ์ บน ย ยักษ์"],
    ["สายสิญจ์", "สายสิญจน์", "มี น หนู การันต์"],
    ["พิธีพิถัน", "พิถีพิถัน", "สระอี อยู่หน้า"],
    ["รายละเอียดละออ", "รายละเอียดลออ", "ละออ"],
    ["ผลัดวันประกันพัก", "ผัดวันประกันพัก", "ผัด (เลื่อนเวลา)"],
    ["ผัดเปลี่ยน", "ผลัดเปลี่ยน", "ผลัด (เปลี่ยนหมุนเวียน)"],
    ["ผัดผ้า", "ผลัดผ้า", "ผลัด (เปลี่ยนผ้า)"],
    ["ผัดใบ", "ผลัดใบ", "ผลัด (ร่วงหลุดใบ)"],
    ["เกษียร", "เกษียณ", "เกษียณ (อายุราชการ)"],
    ["เกษียณสมุทร", "เกษียรสมุทร", "เกษียร (น้ำนม)"],
    ["หลงไหล", "หลงใหล", "สระไอไม้ม้วน"],
    ["ใหลหลง", "ใหลหลง", "สระไอไม้ม้วน"],
    ["น้ำใหล", "น้ำไหล", "สระไอไม้มลาย (ไหล)"],
    ["ลำใย", "ลำไย", "สระไอไม้มลาย"],
    ["กาละเทศะ", "กาลเทศะ", "ไม่มี สระอะ"],
    ["ขมักเขม้น", "ขะมักเขม้น", "มี สระอะ"],
    ["ขะมุกขะมัว", "ขมุกขมัว", "ไม่มี สระอะ"],
    ["คริสมาสต์", "คริสต์มาส", "ต การันต์ อยู่ที่ คริสต์"],
    ["บิณฑบาตร", "บิณฑบาต", "ไม่มี ร การันต์"],
    ["ตักบาตร", "ตักบาตร", "มี ร การันต์"],
    ["บาตรพระ", "บาตร", "มี ร การันต์"],
    ["ผาสุข", "ผาสุก", "ก ไก่ สะกด"],
    ["เกษร", "เกสร", "ส เสือ สะกด"],
    ["ช็อคโกแลต", "ช็อกโกแลต", "ใช้ไม้ไต่คู้"],
    ["ไอศครีม", "ไอศกรีม", "ก ไก่ สะกด"],
    ["คุ๊กกี้", "คุกกี้", "ไม่มีไม้ตรี"],
    ["ปรากฎ", "ปรากฏ", "ใช้ ฏ ฎัก"],
    ["กตหมาย", "กฎหมาย", "ใช้ ฎ ชฎา"],
    ["มงกุฏ", "มงกุฎ", "ใช้ ฎ ชฎา"],
    ["กระรัต", "กะรัต", "ไม่มี ร ควบ"],
    ["กราฝาก", "กาฝาก", "ไม่มี ร ควบ"],
    ["กำมหยี่", "กำมะหยี่", "มี สระอะ"],
    ["ดำลง", "ดำรง", "ร รื่น"],
    ["ดำลงตำแหน่ง", "ดำรงตำแหน่ง", "ร รื่น"],
    ["ศักศรี", "ศักดิ์ศรี", "มี ด การันต์"],
    ["ศักดิ์สิทธิ", "ศักดิ์สิทธิ์", "มี ์ บน ธ ธง"],
    ["สิทธิ์มนุษยชน", "สิทธิมนุษยชน", "ไม่มี ์ บน ธิ"],
    ["มนทล", "มณฑล", "ณ ณรงค์"],
    ["ญาต", "ญาติ", "มี สระอิ"],
    ["สัญชาติญาณ", "สัญชาตญาณ", "ไม่มี สระอิ"],
    ["ภูมปัญญา", "ภูมิปัญญา", "มี สระอิ"],
    ["กิตติมศักดิ", "กิตติมศักดิ์", "มี การันต์"],
    ["อุตสาหกรรม", "อุตสาหกรรม", "อุตสาหกรรม"],
    ["นวัฒกรรม", "นวัตกรรม", "ต ตะเฒ่า/ต เต่า"],
    ["ปฎิสัมพันธ์", "ปฏิสัมพันธ์", "ใช้ ฏ ฎัก"],
    ["ปฎิกิริยา", "ปฏิกิริยา", "ใช้ ฏ ฎัก"],
    ["กฎเกนท์", "กฎเกณฑ์", "ใช้ ฑ ฒิ"],
    ["ดีกา", "ฎีกา", "ใช้ ฎ ชฎา"],
    ["คำพังเผย", "คำพังเพย", "ย ยักษ์ สะกด"],
    ["อุประวรรค", "อุปสรรค", "อุปสรรค"],
    ["ทศวรรศ", "ทศวรรษ", "ษ ฤๅษี สะกด"],
    ["คริสต์ศาสนา", "คริสตศาสนา", "ไม่มี ์"],
    ["พุธศาสนา", "พุทธศาสนา", "พุทธ"],
    ["อธิฐาน", "อธิษฐาน", "ษ ฤๅษี"],
    ["อานิสงค์", "อานิสงส์", "ส เสือ การันต์"],
    ["สวรรคต", "สวรรคต", "สวรรคต"],
    ["สังเกตุการณ์", "สังเกตการณ์", "สังเกต (ไม่มี สระอุ)"],
    ["ประสบการณ์", "ประสบการณ์", "ประสบการณ์"],
    ["สาเหตุ", "สาเหตุ", "สาเหตุ"],
    ["เบญจเพศ", "เบญจเพส", "ส เสือ สะกด"],
    ["อเนจอนาถ", "อเนจอนาถ", "อเนจอนาถ"],
    ["สมเพช", "สมเพช", "ช ช้าง สะกด"],
    ["สมเพชเวทนา", "สมเพชเวทนา", "สมเพช"],
    ["เวทมนตร์", "เวทมนตร์", "ร ท การันต์"],
    ["เวทมนต์", "เวทมนตร์", "ร ท การันต์"],
    ["สวดมนต์", "สวดมนต์", "ต การันต์"],
    ["ศิษย์เก่า", "ศิษย์เก่า", "ย ยักษ์ การันต์"],
    ["ศิษย์เอก", "ศิษย์เอก", "ย ยักษ์ การันต์"],
    ["ปรารถนา", "ปรารถนา", "มี ร ควบ"],
    ["ปาฏิหาริย์", "ปาฏิหาริย์", "ย ยักษ์ การันต์"],
    ["ปาฏิหารย์", "ปาฏิหาริย์", "สระอิ บน ร รื่น"],
    ["อัศจรรย์", "อัศจรรย์", "ย ยักษ์ การันต์"],
    ["อัจฉริยะ", "อัจฉริยะ", "ฉ ฉิ่ง"],
    ["สถาปนิก", "สถาปนิก", "ก ไก่ สะกด"],
    ["มโนภาพ", "มโนภาพ", "มโนภาพ"],
    ["จินตนาการ", "จินตนาการ", "จินตนาการ"],
    ["วิเคราะห์", "วิเคราะห์", "ห การันต์"],
    ["วิเคราห์", "วิเคราะห์", "มี ์"],
    ["สังเคราะห์", "สังเคราะห์", "ห การันต์"],
    ["อนุเคราะห์", "อนุเคราะห์", "ห การันต์"],
    ["สงเคราะห์", "สงเคราะห์", "ห การันต์"],
    ["เคราะห์ร้าย", "เคราะห์ร้าย", "ห การันต์"],
    ["เคราะห์ดี", "เคราะห์ดี", "ห การันต์"],
    ["กะทันหัน", "กะทันหัน", "กะ (ไม่มี ร)"],
    ["กระทันหัน", "กะทันหัน", "กะ (ไม่มี ร)"],
    ["ประณีต", "ประณีต", "ณ ณรงค์ และ ต เต่า"],
    ["ปราณีต", "ประณีต", "ประ (สระอะ)"],
    ["ประณีตบรรจง", "ประณีตบรรจง", "ประณีต"],
    ["กาลเวลา", "กาลเวลา", "ล ลิง"],
    ["กาญจนาภิเษก", "กาญจนาภิเษก", "กาญจนา"],
    ["กุญแจ", "กุญแจ", "ญ หญิง สะกด"],
    ["กุญแจมือ", "กุญแจมือ", "ญ หญิง สะกด"],
    ["สัญญา", "สัญญา", "ญ หญิง"],
    ["สัญญาลักษณ์", "สัญลักษณ์", "สัญลักษณ์"],
    ["สัญญลักษณ์", "สัญลักษณ์", "สัญลักษณ์"],
    ["สัญลักณ์", "สัญลักษณ์", "ษ ฤๅษี และ ก การันต์"],
    ["พากเพียร", "พากเพียร", "ก ไก่ สะกด (พาก)"],
    ["ภาคเพียร", "พากเพียร", "พากเพียร"],
    ["พากภูมิ", "ภาคภูมิ", "ภาคภูมิใจ"],
    ["ภาคภูมิ", "ภาคภูมิ", "ภาคภูมิใจ"],
    ["ภูมิใจ", "ภูมิใจ", "ภูมิใจ"],
    ["ภูมิปัญญา", "ภูมิปัญญา", "ภูมิปัญญา"],
    ["ภูมิอากาศ", "ภูมิอากาศ", "ภูมิอากาศ"],
    ["ภูมิประเทศ", "ภูมิประเทศ", "ภูมิประเทศ"],
    ["กะพริบตา", "กะพริบตา", "กะ (ไม่มี ร)"],
    ["กระพริบตา", "กะพริบตา", "กะ (ไม่มี ร)"]
  ];

  sheet.getRange(2, 1, words.length, 3).setValues(words);
}

function getAssignments(targetClass) {
  var sheet = getOrCreateSheet('Assignments', ['AssignmentID', 'Title', 'Content', 'TargetClass', 'CreatedAt']);
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var id = data[i][0];
    var title = data[i][1];
    var content = data[i][2];
    var cls = data[i][3] || 'ทุกชั้น';
    var createdAt = data[i][4];
    
    if (!targetClass || targetClass === 'all' || cls === 'ทุกชั้น' || cls === targetClass) {
      result.push({
        id: id,
        title: title,
        content: content,
        targetClass: cls,
        createdAt: createdAt,
        rowIdx: i + 1
      });
    }
  }
  return result.reverse();
}

function saveAssignment(p) {
  var sheet = getOrCreateSheet('Assignments', ['AssignmentID', 'Title', 'Content', 'TargetClass', 'CreatedAt']);
  var id = 'ASN_' + Date.now();
  sheet.appendRow([id, p.title || '', p.content || '', p.targetClass || 'ทุกชั้น', new Date()]);
  return { status: 'success', id: id };
}

function deleteAssignment(p) {
  var sheet = getOrCreateSheet('Assignments', ['AssignmentID', 'Title', 'Content', 'TargetClass', 'CreatedAt']);
  if (p.rowIdx && p.rowIdx > 1) {
    sheet.deleteRow(parseInt(p.rowIdx));
    return { status: 'success' };
  }
  if (p.id) {
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == p.id) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
  }
  return { status: 'success' };
}

// --- Homework Topics and Student Homework Edit API ---

function getHwTopics(targetClass) {
  var sheet = getOrCreateSheet('HwTopics', ['TopicID', 'Title', 'TargetClass', 'CreatedAt']);
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var id = data[i][0];
    var title = data[i][1];
    var cls = data[i][2] || 'ทุกชั้น';
    var createdAt = data[i][3];
    
    if (!targetClass || targetClass === 'all' || cls === 'ทุกชั้น' || cls === targetClass) {
      result.push({
        id: id,
        title: title,
        targetClass: cls,
        createdAt: createdAt,
        rowIdx: i + 1
      });
    }
  }
  return result.reverse();
}

function saveHwTopic(p) {
  var sheet = getOrCreateSheet('HwTopics', ['TopicID', 'Title', 'TargetClass', 'CreatedAt']);
  var id = 'HWT_' + Date.now();
  sheet.appendRow([id, p.title || '', p.targetClass || 'ทุกชั้น', new Date()]);
  return { status: 'success', id: id };
}

function deleteHwTopic(p) {
  var sheet = getOrCreateSheet('HwTopics', ['TopicID', 'Title', 'TargetClass', 'CreatedAt']);
  if (p.rowIdx && p.rowIdx > 1) {
    sheet.deleteRow(parseInt(p.rowIdx));
    return { status: 'success' };
  }
  if (p.id) {
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == p.id) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
  }
  return { status: 'success' };
}

function editHomework(p) {
  var sheet = getOrCreateSheet('homework');
  var row = parseInt(p.rowIdx);
  
  sheet.getRange(row, 2).setValue(p.className);
  sheet.getRange(row, 3).setValue(p.name1);
  sheet.getRange(row, 4).setValue(p.name2 || '-');
  sheet.getRange(row, 5).setValue(p.taskName);
  
  if (p.file1) {
    var f1 = uploadFileToDrive(p.file1.data, p.file1.type, p.file1.name);
    sheet.getRange(row, 6).setValue(f1);
  }
  
  if (p.file2) {
    var f2 = uploadFileToDrive(p.file2.data, p.file2.type, p.file2.name);
    sheet.getRange(row, 7).setValue(f2);
  }
  
  sheet.getRange(row, 8).setValue('รอตรวจ');
}

// --- Lessons (ระบบบทเรียน) ---
function getLessons() {
  var sheet = getOrCreateSheet('lessons', ['LessonNo', 'Topic', 'ExternalLink', 'ContentType', 'TargetClass', 'CreatedAt', 'Term']);
  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    list.push({
      rowIdx: i + 1,
      lessonNo: parseInt(data[i][0]),
      topic: data[i][1],
      externalLink: data[i][2],
      contentType: data[i][3],
      targetClass: data[i][4],
      createdAt: data[i][5],
      term: data[i][6] ? parseInt(data[i][6]) : 1
    });
  }
  return list;
}

function saveLesson(p) {
  var sheet = getOrCreateSheet('lessons', ['LessonNo', 'Topic', 'ExternalLink', 'ContentType', 'TargetClass', 'CreatedAt', 'Term']);
  
  // ตรวจสอบและเพิ่มหัวคอลัมน์ Term (คอลัมน์ที่ 7) ถ้ายังไม่มี
  var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headerRow.length < 7 || !headerRow[6] || headerRow[6] === '') {
    sheet.getRange(1, 7).setValue('Term');
    sheet.getRange(1, 7).setFontWeight('bold').setBackground('#e0f2fe');
  }

  var data = sheet.getDataRange().getValues();
  var lessonNo = parseInt(p.lessonNo);
  var targetClass = p.targetClass || 'ทั้งหมด';
  var term = p.term ? parseInt(p.term) : 1;
  var foundRow = -1;
  
  // ค้นหาบทเรียนโดยเช็คทั้ง LessonNo และ Term
  for (var i = 1; i < data.length; i++) {
    var rowLessonNo = parseInt(data[i][0]);
    var rowTerm = data[i][6] ? parseInt(data[i][6]) : 1;
    if (rowLessonNo === lessonNo && rowTerm === term) {
      foundRow = i + 1;
      break;
    }
  }
  
  var dtStr = new Date().toISOString();
  if (foundRow !== -1) {
    sheet.getRange(foundRow, 2).setValue(p.topic);
    sheet.getRange(foundRow, 3).setValue(p.externalLink);
    sheet.getRange(foundRow, 4).setValue(p.contentType || 'link');
    sheet.getRange(foundRow, 5).setValue(targetClass);
    sheet.getRange(foundRow, 6).setValue(dtStr);
    sheet.getRange(foundRow, 7).setValue(term);
  } else {
    sheet.appendRow([lessonNo, p.topic, p.externalLink, p.contentType || 'link', targetClass, dtStr, term]);
  }
  return { status: "success", message: "บันทึกบทเรียนสำเร็จ" };
}

// =========================================================================
// 📝 12. ส่วนระบบสอบออนไลน์ (Online Examination System)
// =========================================================================
function getExams() {
  var sheet = getOrCreateSheet('exams');
  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) {
      var ansKey = [];
      try {
        ansKey = typeof data[i][7] === 'string' ? JSON.parse(data[i][7]) : (Array.isArray(data[i][7]) ? data[i][7] : []);
      } catch (e) {
        ansKey = String(data[i][7] || '').split(',');
      }
      list.push({
        rowIdx: i + 1,
        semester: String(data[i][0]),
        grade: String(data[i][1]),
        subject: String(data[i][2]),
        totalQuestions: parseInt(data[i][3]) || 10,
        choicesCount: parseInt(data[i][4]) || 3,
        timeLimit: parseInt(data[i][5]) || 60,
        pdfId: String(data[i][6] || ''),
        answerKey: ansKey
      });
    }
  }
  return list;
}

function saveExamConfig(p) {
  var sheet = getOrCreateSheet('exams');
  var data = sheet.getDataRange().getValues();
  var ansKeyJson = JSON.stringify(p.answerKey || []);
  var finalPdfId = p.pdfId || '';

  // หากมีการแนบไฟล์ PDF มาพร้อมกับ config
  if (p.pdfData || (p.pdfFile && p.pdfFile.data)) {
    var rawB64 = p.pdfData || p.pdfFile.data;
    var filename = p.filename || (p.pdfFile ? p.pdfFile.name : '') || ('exam_' + p.grade + '_' + p.subject + '_' + Date.now() + '.pdf');
    var mime = p.mimeType || (p.pdfFile ? p.pdfFile.type : '') || 'application/pdf';
    var uploaded = uploadExamPDF({ data: rawB64, filename: filename, mimeType: mime });
    if (uploaded && uploaded.fileId) {
      finalPdfId = uploaded.fileId;
    }
  }

  var foundRow = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(p.semester) && String(data[i][1]) === String(p.grade) && String(data[i][2]) === String(p.subject)) {
      foundRow = i + 1;
      break;
    }
  }
  if (foundRow > -1) {
    sheet.getRange(foundRow, 4, 1, 5).setValues([[
      parseInt(p.totalQuestions) || 10,
      parseInt(p.choicesCount) || 4,
      parseInt(p.timeLimit) || 60,
      finalPdfId,
      ansKeyJson
    ]]);
  } else {
    sheet.appendRow([
      String(p.semester),
      String(p.grade),
      String(p.subject),
      parseInt(p.totalQuestions) || 10,
      parseInt(p.choicesCount) || 4,
      parseInt(p.timeLimit) || 60,
      finalPdfId,
      ansKeyJson
    ]);
  }
  return { status: "success", message: "บันทึกข้อสอบเรียบร้อย", pdfId: finalPdfId };
}

function uploadExamPDF(p) {
  try {
    if (!p.data) return { status: "error", message: "ไม่มีข้อมูลไฟล์" };
    var examFolderId = "1XfCSFFHFZDTp5s7EcNelMNfifCHLHziU";
    var folder;
    try {
      folder = DriveApp.getFolderById(examFolderId);
    } catch(fErr1) {
      try {
        folder = DriveApp.getFolderById(FOLDER_ID);
      } catch(fErr2) {
        folder = DriveApp.getRootFolder();
      }
    }
    var cleanData = p.data;
    if (cleanData.indexOf(',') !== -1) {
      cleanData = cleanData.split(',')[1];
    }
    var bytes = Utilities.base64Decode(cleanData);
    var filename = p.filename || ('exam_' + Date.now() + '.pdf');
    var blob = Utilities.newBlob(bytes, p.mimeType || 'application/pdf', filename);
    var file = folder.createFile(blob);
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch(shErr){}
    var fileId = file.getId();
    return { status: "success", fileId: fileId, url: file.getUrl() };
  } catch(err) {
    return { status: "error", message: err.toString() };
  }
}

function getExamScores() {
  var sheet = getOrCreateSheet('exam_scores');
  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][1]) {
      list.push({
        rowIdx: i + 1,
        date: data[i][0],
        grade: String(data[i][1]),
        name: String(data[i][2]),
        semester: String(data[i][3]),
        subject: String(data[i][4]),
        score: parseInt(data[i][5]) || 0,
        totalQuestions: parseInt(data[i][6]) || 10,
        attemptCount: parseInt(data[i][7]) || 1
      });
    }
  }
  return list.reverse();
}

function saveExamScore(p) {
  var sheet = getOrCreateSheet('exam_scores');
  var data = sheet.getDataRange().getValues();
  var dtStr = new Date().toISOString();
  var attempt = 1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(p.grade) && String(data[i][2]) === String(p.name) && String(data[i][3]) === String(p.semester) && String(data[i][4]) === String(p.subject)) {
      attempt++;
    }
  }
  sheet.appendRow([
    dtStr,
    String(p.grade),
    String(p.name),
    String(p.semester),
    String(p.subject),
    parseInt(p.score) || 0,
    parseInt(p.totalQuestions) || 10,
    attempt
  ]);
  return { status: "success", attemptCount: attempt };
}