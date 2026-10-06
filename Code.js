// ==========================================
// Google Apps Script: Code.js
// ระบบปฏิทินกิจกรรมการศึกษา (Thai Buddhist Era + Hijri Calendar)
// ผูกกับ Google Sheet: 1Yz5oBkfpb8ERqojxhBkzkL1mxRIvK4BESrerGcKPFt4
// ==========================================

var SPREADSHEET_ID = '1Yz5oBkfpb8ERqojxhBkzkL1mxRIvK4BESrerGcKPFt4';

function getSpreadsheet() {
  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (e) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) return ss;
    throw new Error('ไม่สามารถเข้าถึง Google Sheet ID: ' + SPREADSHEET_ID + ' (' + e.message + ')');
  }
}

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('ระบบปฏิทินกิจกรรมประจำปีการศึกษา (พ.ศ. / ฮ.ศ.)')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * เตรียมชีตและข้อมูลเริ่มต้นอัตโนมัติหากยังไม่มีชีต
 */
function getOrCreateSheets() {
  var ss = getSpreadsheet();
  
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
    eventSheet.appendRow(['id', 'title', 'departmentId', 'startDate', 'endDate', 'academicYear', 'semester', 'location', 'description', 'createdBy', 'creatorName', 'creatorDept', 'createdAt', 'updatedAt']);
    
    // ใส่กิจกรรมตั้งต้น (วันหยุด พ.ศ. / ฮ.ศ.)
    var defaults = getDefaultEvents();
    for (var i = 0; i < defaults.length; i++) {
      var d = defaults[i];
      eventSheet.appendRow([
        d.id, d.title, d.departmentId, d.startDate, d.endDate,
        d.academicYear, d.semester, d.location || '', d.description || '',
        'ระบบ', 'ส่วนกลาง', 'ฝ่ายบริหารทั่วไปและแผนงาน', new Date().toISOString(), new Date().toISOString()
      ]);
    }
  }

  // 3. ชีตผู้ใช้งานและสมาชิก (Users)
  var userSheet = ss.getSheetByName('Users');
  if (!userSheet) {
    userSheet = ss.insertSheet('Users');
    userSheet.appendRow(['id', 'name', 'department', 'email', 'phone', 'role', 'createdAt']);
    userSheet.appendRow(['u_admin', 'ผู้ดูแลระบบกลาง', 'ฝ่ายบริหารทั่วไปและแผนงาน', 'admin@school.ac.th', '0812345678', 'admin', new Date().toISOString()]);
  }

  // 4. ชีตการตั้งค่า (Settings)
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
 * ดึงข้อมูลตั้งต้นทั้งหมดเพื่อส่งให้หน้าเว็บ (Init Data)
 */
function getInitialData() {
  try {
    getOrCreateSheets();
    return {
      success: true,
      departments: getDepartments(),
      events: getEvents(),
      users: getUsers(),
      settings: getSettings()
    };
  } catch (err) {
    return {
      success: false,
      error: err.message,
      departments: getDefaultDepartments(),
      events: getDefaultEvents(),
      users: [],
      settings: { adminPinSet: true, currentAcademicYear: '2568' }
    };
  }
}

function getDefaultDepartments() {
  return [
    { id: 'dept_holiday', name: 'วันหยุดราชการ / สำคัญ', color: '#dc2626' },
    { id: 'dept_acad', name: 'ฝ่ายวิชาการ', color: '#2563eb' },
    { id: 'dept_student', name: 'ฝ่ายกิจการนักเรียน', color: '#16a34a' },
    { id: 'dept_admin', name: 'ฝ่ายบริหารทั่วไปและแผนงาน', color: '#ea580c' },
    { id: 'dept_finance', name: 'ฝ่ายงบประมาณและการเงิน', color: '#9333ea' }
  ];
}

function getDefaultEvents() {
  return [
    {
      id: 'h_newyear_2025',
      title: 'วันขึ้นปีใหม่',
      departmentId: 'dept_holiday',
      startDate: '2025-01-01',
      endDate: '2025-01-01',
      academicYear: '2567',
      semester: '2',
      location: '-',
      description: 'วันหยุดราชการสากล',
      creatorName: 'ระบบ',
      creatorDept: 'ส่วนกลาง'
    },
    {
      id: 'h_eid_fitr_2568',
      title: 'วันตรุษอีดิ้ลฟิฏริ (ฮ.ศ. 1446)',
      departmentId: 'dept_holiday',
      startDate: '2025-03-31',
      endDate: '2025-04-01',
      academicYear: '2567',
      semester: '2',
      location: '-',
      description: 'วันเฉลิมฉลองออกบวช วันสำคัญทางศาสนาอิสลาม (1 เชาวาล 1446)',
      creatorName: 'ระบบ',
      creatorDept: 'ส่วนกลาง'
    },
    {
      id: 'h_songkran_2025',
      title: 'วันสงกรานต์',
      departmentId: 'dept_holiday',
      startDate: '2025-04-13',
      endDate: '2025-04-15',
      academicYear: '2567',
      semester: 'ปิดภาคเรียน',
      location: '-',
      description: 'วันหยุดราชการและวันขึ้นปีใหม่ไทย',
      creatorName: 'ระบบ',
      creatorDept: 'ส่วนกลาง'
    },
    {
      id: 'h_eid_adha_2568',
      title: 'วันตรุษอีดิ้ลอัฎฮา (ฮ.ศ. 1446)',
      departmentId: 'dept_holiday',
      startDate: '2025-06-06',
      endDate: '2025-06-08',
      academicYear: '2568',
      semester: '1',
      location: '-',
      description: 'วันฉลองเชือดสัตว์พลีทาน (10 ซุลฮิจญะฮ์ 1446)',
      creatorName: 'ระบบ',
      creatorDept: 'ส่วนกลาง'
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
      creatorName: 'ระบบ',
      creatorDept: 'ส่วนกลาง'
    }
  ];
}

// ------------------------------
// API: Users (ระบบสมาชิกและสังกัดฝ่าย)
// ------------------------------
function getUsers() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Users');
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  var results = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[0]) {
      results.push({
        id: String(row[0]),
        name: String(row[1] || ''),
        department: String(row[2] || ''),
        email: String(row[3] || ''),
        phone: String(row[4] || ''),
        role: String(row[5] || 'staff')
      });
    }
  }
  return results;
}

function registerUser(userData) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Users');
  if (!sheet) {
    getOrCreateSheets();
    sheet = ss.getSheetByName('Users');
  }

  var id = userData.id || ('u_' + new Date().getTime());
  var role = userData.role || 'staff';
  var now = new Date().toISOString();

  sheet.appendRow([
    id,
    userData.name || 'ไม่ระบุชื่อ',
    userData.department || 'ไม่ระบุหน่วยงาน',
    userData.email || '',
    userData.phone || '',
    role,
    now
  ]);

  return {
    success: true,
    user: {
      id: id,
      name: userData.name,
      department: userData.department,
      email: userData.email,
      phone: userData.phone,
      role: role
    }
  };
}

function deleteUser(userId, adminPin) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Users');
  if (!sheet) return { success: false, message: 'ไม่พบชีต Users' };

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(userId)) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, message: 'ไม่พบผู้ใช้ที่ต้องการลบ' };
}

// ------------------------------
// API: Departments
// ------------------------------
function getDepartments() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Departments');
  if (!sheet) return getDefaultDepartments();

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

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
  var ss = getSpreadsheet();
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
    sheet.getRange(rowIndex, 2).setValue(deptData.name);
    sheet.getRange(rowIndex, 3).setValue(deptData.color);
  } else {
    sheet.appendRow([id, deptData.name, deptData.color, new Date().toISOString()]);
  }

  return { success: true, id: id };
}

function deleteDepartment(deptId, adminPin) {
  var ss = getSpreadsheet();
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
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Events');
  if (!sheet) return getDefaultEvents();

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
        createdBy: String(row[9] || ''),
        creatorName: String(row[10] || ''),
        creatorDept: String(row[11] || '')
      });
    }
  }
  return results;
}

function saveEvent(eventData) {
  var ss = getSpreadsheet();
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
    sheet.getRange(rowIndex, 2).setValue(eventData.title);
    sheet.getRange(rowIndex, 3).setValue(eventData.departmentId);
    sheet.getRange(rowIndex, 4).setValue(eventData.startDate);
    sheet.getRange(rowIndex, 5).setValue(eventData.endDate);
    sheet.getRange(rowIndex, 6).setValue(eventData.academicYear);
    sheet.getRange(rowIndex, 7).setValue(eventData.semester);
    sheet.getRange(rowIndex, 8).setValue(eventData.location);
    sheet.getRange(rowIndex, 9).setValue(eventData.description);
    sheet.getRange(rowIndex, 11).setValue(eventData.creatorName || '');
    sheet.getRange(rowIndex, 12).setValue(eventData.creatorDept || '');
    sheet.getRange(rowIndex, 14).setValue(now);
  } else {
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
      eventData.creatorName || 'ไม่ระบุชื่อ',
      eventData.creatorDept || 'ไม่ระบุฝ่าย',
      now,
      now
    ]);
  }

  return { success: true, id: id };
}

function deleteEvent(eventId, adminPin) {
  var ss = getSpreadsheet();
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
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Settings');
  if (!sheet) return pin === '1234';

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === 'admin_pin') {
      return String(data[i][1]) === String(pin);
    }
  }
  return pin === '1234';
}

function getSettings() {
  var ss = getSpreadsheet();
  var settings = { currentAcademicYear: '2568' };
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
