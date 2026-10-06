// ==========================================
// Google Apps Script: Code.js
// ระบบปฏิทินกิจกรรมการศึกษา (Thai Buddhist Era + Hijri Calendar)
// ==========================================

function doGet(e) {
  // ให้สิทธิ์การเข้าถึงแบบ Web App
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('ระบบปฏิทินกิจกรรมประจำปีการศึกษา')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * ฟังก์ชันสร้าง/เตรียมชีตอัตโนมัติหากยังไม่มี
 */
function getOrCreateSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    // กรณีที่ไม่ได้ผูกกับ Google Sheets ปัจจุบัน (เช่น ทดสอบแบบ Standalone)
    return null;
  }
  
  // 1. ชีตฝ่าย (Departments)
  var deptSheet = ss.getSheetByName('Departments');
  if (!deptSheet) {
    deptSheet = ss.insertSheet('Departments');
    deptSheet.appendRow(['id', 'name', 'color', 'createdAt']);
    deptSheet.appendRow(['dept_acad', 'ฝ่ายวิชาการ', '#2563eb', new Date().toISOString()]);
    deptSheet.appendRow(['dept_student', 'ฝ่ายกิจการนักเรียน', '#16a34a', new Date().toISOString()]);
    deptSheet.appendRow(['dept_admin', 'ฝ่ายบริหารทั่วไปและแผนงาน', '#ea580c', new Date().toISOString()]);
    deptSheet.appendRow(['dept_finance', 'ฝ่ายงบประมาณและการเงิน', '#9333ea', new Date().toISOString()]);
    deptSheet.appendRow(['dept_holiday', 'วันหยุด/เทศกาล', '#dc2626', new Date().toISOString()]);
  }

  // 2. ชีตกิจกรรม (Events)
  var eventSheet = ss.getSheetByName('Events');
  if (!eventSheet) {
    eventSheet = ss.insertSheet('Events');
    eventSheet.appendRow(['id', 'title', 'departmentId', 'startDate', 'endDate', 'academicYear', 'semester', 'location', 'description', 'createdBy', 'createdAt', 'updatedAt']);
  }

  // 3. ชีตการตั้งค่าระบบ (Settings)
  var settingSheet = ss.getSheetByName('Settings');
  if (!settingSheet) {
    settingSheet = ss.insertSheet('Settings');
    settingSheet.appendRow(['key', 'value']);
    settingSheet.appendRow(['admin_pin', '1234']); // รหัสผ่านแอดมินเริ่มต้น
    settingSheet.appendRow(['current_academic_year', '2568']);
  }

  return ss;
}

/**
 * ดึงข้อมูลเบื้องต้นทั้งหมด (Init Data)
 */
function getInitialData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    return {
      success: true,
      departments: getDefaultDepartments(),
      events: getDefaultEvents(),
      settings: { adminPinSet: true, currentAcademicYear: '2568' }
    };
  }

  getOrCreateSheets();

  var departments = getDepartments();
  var events = getEvents();
  var settings = getSettings();

  return {
    success: true,
    departments: departments,
    events: events,
    settings: settings
  };
}

/**
 * ข้อมูลเริ่มต้นจำลอง (Fallback กรณีทดสอบแบบยังไม่ผูกชีต)
 */
function getDefaultDepartments() {
  return [
    { id: 'dept_acad', name: 'ฝ่ายวิชาการ', color: '#2563eb' },
    { id: 'dept_student', name: 'ฝ่ายกิจการนักเรียน', color: '#16a34a' },
    { id: 'dept_admin', name: 'ฝ่ายบริหารทั่วไปและแผนงาน', color: '#ea580c' },
    { id: 'dept_finance', name: 'ฝ่ายงบประมาณและการเงิน', color: '#9333ea' },
    { id: 'dept_holiday', name: 'วันหยุดราชการ / สำคัญ', color: '#dc2626' }
  ];
}

function getDefaultEvents() {
  return [
    {
      id: 'h_newyear',
      title: 'วันขึ้นปีใหม่',
      departmentId: 'dept_holiday',
      startDate: '2025-01-01',
      endDate: '2025-01-01',
      academicYear: '2567',
      semester: '2',
      location: '-',
      description: 'วันหยุดราชการสากล',
      isLockedHoliday: true
    },
    {
      id: 'h_eid_fitr_2568',
      title: 'วันอีดิ้ลฟิฏริ (ฮ.ศ. 1446)',
      departmentId: 'dept_holiday',
      startDate: '2025-03-31',
      endDate: '2025-03-31',
      academicYear: '2567',
      semester: '2',
      location: '-',
      description: 'วันเฉลิมฉลองออกบวช วันสำคัญทางศาสนาอิสลาม (1 เชาวาล 1446)',
      isLockedHoliday: true
    },
    {
      id: 'h_songkran_2568',
      title: 'วันสงกรานต์',
      departmentId: 'dept_holiday',
      startDate: '2025-04-13',
      endDate: '2025-04-15',
      academicYear: '2567',
      semester: 'ปิดภาคเรียน',
      location: '-',
      description: 'วันหยุดราชการและวันขึ้นปีใหม่ไทย',
      isLockedHoliday: true
    },
    {
      id: 'h_eid_adha_2568',
      title: 'วันอีดิ้ลอัฎฮา (ฮ.ศ. 1446)',
      departmentId: 'dept_holiday',
      startDate: '2025-06-06',
      endDate: '2025-06-06',
      academicYear: '2568',
      semester: '1',
      location: '-',
      description: 'วันฉลองเชือดสัตว์พลีทาน (10 ซุลฮิจญะฮ์ 1446)',
      isLockedHoliday: true
    },
    {
      id: 'h_hijri_newyear_1447',
      title: 'วันขึ้นปีใหม่ฮิจเราะห์ศักราช 1447',
      departmentId: 'dept_holiday',
      startDate: '2025-06-26',
      endDate: '2025-06-26',
      academicYear: '2568',
      semester: '1',
      location: '-',
      description: '1 มุฮัรรอม ฮ.ศ. 1447',
      isLockedHoliday: true
    }
  ];
}

// ------------------------------
// API: Departments
// ------------------------------
function getDepartments() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return getDefaultDepartments();

  var sheet = ss.getSheetByName('Departments');
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  var headers = data[0];
  var results = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[0]) {
      results.push({
        id: String(row[0]),
        name: String(row[1]),
        color: String(row[2]) || '#2563eb'
      });
    }
  }
  return results;
}

function saveDepartment(deptData, adminPin) {
  if (!verifyAdminPin(adminPin)) {
    throw new Error('รหัสผ่านผู้ดูแลระบบ (Admin PIN) ไม่ถูกต้อง');
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Departments');
  var data = sheet.getDataRange().getValues();
  var id = deptData.id || ('dept_' + new Date().getTime());

  var rowIndex = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex > -1) {
    // อัปเดตฝ่ายเดิม
    sheet.getRange(rowIndex, 2).setValue(deptData.name);
    sheet.getRange(rowIndex, 3).setValue(deptData.color);
  } else {
    // เพิ่มฝ่ายใหม่
    sheet.appendRow([id, deptData.name, deptData.color, new Date().toISOString()]);
  }

  return { success: true, id: id };
}

function deleteDepartment(deptId, adminPin) {
  if (!verifyAdminPin(adminPin)) {
    throw new Error('รหัสผ่านผู้ดูแลระบบ (Admin PIN) ไม่ถูกต้อง');
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Departments');
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(deptId)) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, message: 'ไม่พบฝ่ายที่ต้องการลบ' };
}

// ------------------------------
// API: Events
// ------------------------------
function getEvents() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return getDefaultEvents();

  var sheet = ss.getSheetByName('Events');
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  var results = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[0]) {
      results.push({
        id: String(row[0]),
        title: String(row[1]),
        departmentId: String(row[2]),
        startDate: formatDateString(row[3]),
        endDate: formatDateString(row[4] || row[3]),
        academicYear: String(row[5] || ''),
        semester: String(row[6] || ''),
        location: String(row[7] || ''),
        description: String(row[8] || ''),
        createdBy: String(row[9] || '')
      });
    }
  }
  return results;
}

function saveEvent(eventData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Events');
  var data = sheet.getDataRange().getValues();
  var id = eventData.id || ('evt_' + new Date().getTime());

  var rowIndex = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      rowIndex = i + 1;
      break;
    }
  }

  var now = new Date().toISOString();
  if (rowIndex > -1) {
    // แก้ไขกิจกรรม
    sheet.getRange(rowIndex, 2).setValue(eventData.title);
    sheet.getRange(rowIndex, 3).setValue(eventData.departmentId);
    sheet.getRange(rowIndex, 4).setValue(eventData.startDate);
    sheet.getRange(rowIndex, 5).setValue(eventData.endDate);
    sheet.getRange(rowIndex, 6).setValue(eventData.academicYear);
    sheet.getRange(rowIndex, 7).setValue(eventData.semester);
    sheet.getRange(rowIndex, 8).setValue(eventData.location);
    sheet.getRange(rowIndex, 9).setValue(eventData.description);
    sheet.getRange(rowIndex, 12).setValue(now);
  } else {
    // เพิ่มกิจกรรมใหม่
    sheet.appendRow([
      id,
      eventData.title,
      eventData.departmentId,
      eventData.startDate,
      eventData.endDate,
      eventData.academicYear,
      eventData.semester,
      eventData.location || '',
      eventData.description || '',
      eventData.createdBy || 'ทั่วไป',
      now,
      now
    ]);
  }

  return { success: true, id: id };
}

function deleteEvent(eventId, adminPin) {
  // แอดมินสามารถลบกิจกรรมใดก็ได้
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Events');
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(eventId)) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, message: 'ไม่พบกิจกรรมที่ต้องการลบ' };
}

// ------------------------------
// Helpers & Settings
// ------------------------------
function verifyAdminPin(pin) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return pin === '1234';

  var sheet = ss.getSheetByName('Settings');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === 'admin_pin') {
      return String(data[i][1]) === String(pin);
    }
  }
  return pin === '1234'; // ค่าเริ่มต้น
}

function getSettings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settings = { currentAcademicYear: '2568' };
  if (!ss) return settings;

  var sheet = ss.getSheetByName('Settings');
  if (!sheet) return settings;

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === 'current_academic_year') {
      settings.currentAcademicYear = String(data[i][1]);
    }
  }
  return settings;
}

function formatDateString(val) {
  if (!val) return '';
  if (val instanceof Date) {
    var y = val.getFullYear();
    var m = ('0' + (val.getMonth() + 1)).slice(-2);
    var d = ('0' + val.getDate()).slice(-2);
    return y + '-' + m + '-' + d;
  }
  return String(val).split('T')[0];
}
