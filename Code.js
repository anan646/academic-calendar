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
  // รองรับ JSON API ทั้งแบบ GET ธรรมดา และเมื่อ Redirect มาจาก POST
  if (e && e.parameter && e.parameter.action) {
    return handleApiRequest(e, 'GET');
  }

  // หากเรียกเข้ามาโดยไม่มี action ให้ตอบกลับสถานะ JSON API พร้อมใช้งานเสมอ (ป้องกันไม่ให้พังจากการหา index.html ไม่เจอ)
  return ContentService.createTextOutput(JSON.stringify({
    success: true,
    status: 'online',
    message: 'Google Apps Script Academic Calendar API is active and ready.',
    timestamp: new Date().toISOString()
  })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  return handleApiRequest(e, 'POST');
}

/**
 * ตัวประมวลผลคำสั่ง API รวม (Unified API Dispatcher)
 * รองรับทั้ง GET และ POST เพื่อความเสถียรสูงสุดบนทุกอุปกรณ์และเบราว์เซอร์
 */
function handleApiRequest(e, method) {
  var result = { success: false, message: 'Invalid action' };
  var callback = (e && e.parameter && e.parameter.callback) || null;

  try {
    getOrCreateSheets();

    // รวบรวมข้อมูล Parameter จากทั้ง Query String และ POST Body
    var params = {};
    if (e && e.parameter) {
      for (var key in e.parameter) {
        params[key] = e.parameter[key];
      }
    }

    if (e && e.postData && e.postData.contents) {
      try {
        var parsed = JSON.parse(e.postData.contents);
        for (var pKey in parsed) {
          params[pKey] = parsed[pKey];
        }
      } catch (errJson) {
        // อาจเป็น plain form data
      }
    }

    var action = params.action;

    // ถอดรหัส JSON ย่อยหากส่งมาในรูป string (กรณีส่งผ่าน URL Query params)
    ['eventData', 'deptData', 'userData', 'adminData', 'eventsList'].forEach(function(field) {
      if (typeof params[field] === 'string' && params[field].trim().startsWith('{')) {
        try { params[field] = JSON.parse(params[field]); } catch (errParse) {}
      } else if (typeof params[field] === 'string' && params[field].trim().startsWith('[')) {
        try { params[field] = JSON.parse(params[field]); } catch (errParse) {}
      }
    });

    if (action === 'getInitialData') {
      result = getInitialData();
    } else if (action === 'getEvents') {
      result = { success: true, events: getEvents() };
    } else if (action === 'getDepartments') {
      result = { success: true, departments: getDepartments() };
    } else if (action === 'getUsers') {
      result = { success: true, users: getUsers() };
    } else if (action === 'getAdmins') {
      result = { success: true, admins: getAdmins() };
    } else if (action === 'saveEvent') {
      result = saveEvent(params.eventData || params);
    } else if (action === 'batchSaveEvents') {
      result = batchSaveEvents(params.eventsList || []);
    } else if (action === 'deleteEvent') {
      result = deleteEvent(params.id, params.adminPin);
    } else if (action === 'saveDepartment') {
      result = saveDepartment(params.deptData || params, params.adminPin);
    } else if (action === 'deleteDepartment') {
      result = deleteDepartment(params.id, params.adminPin);
    } else if (action === 'registerUser') {
      result = registerUser(params.userData || params);
    } else if (action === 'deleteUser') {
      result = deleteUser(params.userId, params.adminPin);
    } else if (action === 'saveAdmin') {
      result = saveAdmin(params.adminData || params, params.adminPin);
    } else if (action === 'deleteAdmin') {
      result = deleteAdmin(params.adminId, params.adminPin);
    } else if (action === 'changeAdminPin') {
      result = changeAdminPin(params.adminId, params.newPin, params.adminPin);
    } else {
      result = { success: false, message: 'Unrecognized action: ' + action };
    }
  } catch (err) {
    result = { success: false, error: err.message || String(err) };
  }

  var outputText = JSON.stringify(result);
  if (callback) {
    // JSONP Support สำหรับเบราว์เซอร์เก่าหรือเครือข่ายที่มี CORS ข้อจำกัด
    return ContentService.createTextOutput(callback + '(' + outputText + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService.createTextOutput(outputText)
    .setMimeType(ContentService.MimeType.JSON);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * เตรียมชีตและโครงสร้างข้อมูลเริ่มต้นอัตโนมัติหากยังไม่มีชีต
 */
function getOrCreateSheets() {
  var ss = getSpreadsheet();
  
  // 1. ชีตฝ่าย (Departments)
  var deptSheet = ss.getSheetByName('Departments');
  if (!deptSheet) {
    deptSheet = ss.insertSheet('Departments');
    deptSheet.appendRow(['id', 'name', 'color', 'pin', 'createdAt']);
    deptSheet.appendRow(['dept_acad', 'ฝ่ายวิชาการ', '#2563eb', '1111', new Date().toISOString()]);
    deptSheet.appendRow(['dept_student', 'ฝ่ายกิจการนักเรียน', '#16a34a', '2222', new Date().toISOString()]);
    deptSheet.appendRow(['dept_admin', 'ฝ่ายบริหารทั่วไปและแผนงาน', '#ea580c', '3333', new Date().toISOString()]);
    deptSheet.appendRow(['dept_finance', 'ฝ่ายงบประมาณและการเงิน', '#9333ea', '4444', new Date().toISOString()]);
    deptSheet.appendRow(['dept_holiday', 'วันหยุด/เทศกาล', '#dc2626', '', new Date().toISOString()]);
  } else {
    // ตรวจสอบว่ามีคอลัมน์ pin ใน header หรือไม่
    var dHeaders = deptSheet.getRange(1, 1, 1, Math.max(1, deptSheet.getLastColumn())).getValues()[0] || [];
    var hasPin = dHeaders.some(function(h) { return String(h).toLowerCase() === 'pin'; });
    if (!hasPin) {
      deptSheet.getRange(1, dHeaders.length + 1).setValue('pin');
    }
  }

  // 2. ชีตกิจกรรม (Events)
  var eventSheet = ss.getSheetByName('Events');
  if (!eventSheet) {
    eventSheet = ss.insertSheet('Events');
    eventSheet.appendRow(['id', 'title', 'departmentId', 'startDate', 'endDate', 'academicYear', 'semester', 'location', 'description', 'createdBy', 'creatorName', 'creatorDept', 'createdAt', 'updatedAt']);
    
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
    settingSheet.appendRow(['admin_pin', '1234']);
    settingSheet.appendRow(['current_academic_year', '2568']);
  }

  // 5. ชีตผู้ดูแลระบบ (Admins - แอดมินหลัก & แอดมินรอง)
  var adminSheet = ss.getSheetByName('Admins');
  if (!adminSheet) {
    adminSheet = ss.insertSheet('Admins');
    adminSheet.appendRow(['id', 'name', 'department', 'contact', 'role', 'pin', 'createdAt', 'updatedAt']);
    adminSheet.appendRow([
      'admin_primary',
      'แอดมินหลัก (ผู้ดูแลระบบหลัก)',
      'ผู้ดูแลระบบส่วนกลาง',
      'admin@school.ac.th',
      'primary',
      '1234',
      '2568-01-01',
      new Date().toISOString()
    ]);
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
      admins: getAdmins(),
      settings: getSettings()
    };
  } catch (err) {
    return {
      success: false,
      error: err.message,
      departments: getDefaultDepartments(),
      events: getDefaultEvents(),
      users: [],
      admins: getDefaultAdmins(),
      settings: { adminPinSet: true, currentAcademicYear: '2568' }
    };
  }
}

function getDefaultDepartments() {
  return [
    { id: 'dept_ftu', name: 'ปฏิทิน ม.ฟาฏอนี', color: '#047857', pin: '5555' },
    { id: 'dept_holiday', name: 'วันหยุดราชการ / สำคัญ', color: '#dc2626', pin: '' },
    { id: 'dept_acad', name: 'ฝ่ายวิชาการ', color: '#2563eb', pin: '1111' },
    { id: 'dept_student', name: 'ฝ่ายกิจการนักเรียน', color: '#16a34a', pin: '2222' },
    { id: 'dept_admin', name: 'ฝ่ายบริหารทั่วไปและแผนงาน', color: '#ea580c', pin: '3333' },
    { id: 'dept_finance', name: 'ฝ่ายงบประมาณและการเงิน', color: '#9333ea', pin: '4444' }
  ];
}

function getDefaultAdmins() {
  return [
    {
      id: 'admin_primary',
      name: 'แอดมินหลัก (ผู้ดูแลระบบหลัก)',
      department: 'ผู้ดูแลระบบส่วนกลาง',
      contact: 'admin@school.ac.th',
      role: 'primary',
      pin: '1234',
      createdAt: '2568-01-01'
    }
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
      description: 'วันขึ้น 1 มุฮัรร็อม 1447 ศักราชใหม่แห่งโลกอิสลาม',
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
    userData.email || userData.contact || '',
    userData.phone || '',
    role,
    now
  ]);

  return {
    success: true,
    id: id,
    message: 'ลงทะเบียนผู้ใช้งานเรียบร้อยแล้ว'
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
      return { success: true, message: 'ลบผู้ใช้เรียบร้อยแล้ว' };
    }
  }
  return { success: false, message: 'ไม่พบผู้ใช้ที่ต้องการลบ' };
}

// ------------------------------
// API: Departments (ฝ่ายงานและรหัส PIN)
// ------------------------------
function getDepartments() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Departments');
  if (!sheet) return getDefaultDepartments();

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return getDefaultDepartments();

  var headers = data[0] || [];
  var pinColIdx = -1;
  for (var h = 0; h < headers.length; h++) {
    if (String(headers[h]).toLowerCase() === 'pin') {
      pinColIdx = h;
      break;
    }
  }

  var defaults = getDefaultDepartments();
  var defaultPinMap = {};
  defaults.forEach(function(d) {
    if (d.pin) defaultPinMap[d.id] = d.pin;
  });

  var results = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[0]) {
      var dId = String(row[0]);
      var dPin = '';
      if (pinColIdx > -1 && row[pinColIdx] !== undefined && row[pinColIdx] !== null && String(row[pinColIdx]).trim() !== '') {
        dPin = String(row[pinColIdx]).trim();
      } else if (defaultPinMap[dId]) {
        dPin = defaultPinMap[dId];
      }

      results.push({
        id: dId,
        name: String(row[1]),
        color: String(row[2]) || '#2563eb',
        pin: dPin
      });
    }
  }
  return results;
}

function saveDepartment(deptData, adminPin) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Departments');
  if (!sheet) {
    getOrCreateSheets();
    sheet = ss.getSheetByName('Departments');
  }

  var data = sheet.getDataRange().getValues();
  var id = deptData.id || ('dept_' + new Date().getTime());

  var headers = data[0] || [];
  var pinColIdx = -1;
  for (var h = 0; h < headers.length; h++) {
    if (String(headers[h]).toLowerCase() === 'pin') {
      pinColIdx = h;
      break;
    }
  }

  if (pinColIdx === -1) {
    pinColIdx = headers.length;
    sheet.getRange(1, pinColIdx + 1).setValue('pin');
  }

  var rowIndex = -1;
  for (var i = 1; i < data.length; i++) {
    var rowId = String(data[i][0] || '').trim();
    var rowName = String(data[i][1] || '').trim();
    if ((id && rowId === String(id).trim()) || (deptData.name && rowName === String(deptData.name).trim())) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex > -1) {
    if (deptData.name) sheet.getRange(rowIndex, 2).setValue(deptData.name);
    if (deptData.color) sheet.getRange(rowIndex, 3).setValue(deptData.color);
    if (deptData.pin !== undefined && deptData.pin !== null) {
      sheet.getRange(rowIndex, pinColIdx + 1).setValue(String(deptData.pin));
    }
  } else {
    var newRow = [id, deptData.name, deptData.color || '#7c3aed'];
    while (newRow.length < pinColIdx) {
      newRow.push('');
    }
    newRow[pinColIdx] = String(deptData.pin || '');
    newRow.push(new Date().toISOString());
    sheet.appendRow(newRow);
  }

  SpreadsheetApp.flush();
  return { success: true, id: id };
}

function deleteDepartment(deptId, adminPin) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Departments');
  if (!sheet) return { success: false, message: 'ไม่พบชีต Departments' };

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(deptId)) {
      sheet.deleteRow(i + 1);
      SpreadsheetApp.flush();
      return { success: true };
    }
  }
  return { success: false, message: 'ไม่พบฝ่ายที่ต้องการลบ' };
}

// ------------------------------
// API: Admins (แอดมินหลัก & แอดมินรอง)
// ------------------------------
function getAdmins() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Admins');
  if (!sheet) {
    getOrCreateSheets();
    sheet = ss.getSheetByName('Admins');
  }

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return getDefaultAdmins();

  var results = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[0]) {
      results.push({
        id: String(row[0]),
        name: String(row[1] || 'ผู้ดูแลระบบ'),
        department: String(row[2] || 'ผู้ดูแลระบบส่วนกลาง'),
        contact: String(row[3] || ''),
        role: String(row[4] || 'sub_admin'),
        pin: String(row[5] || '1234'),
        createdAt: formatDateString(row[6]) || '2568-01-01'
      });
    }
  }

  // หากไม่มี Primary Admin ในรายการ ให้เพิ่มเริ่มต้นเข้าไป
  var hasPrimary = results.some(function(a) { return a.role === 'primary'; });
  if (!hasPrimary) {
    results.unshift(getDefaultAdmins()[0]);
  }

  return results;
}

function saveAdmin(adminData, adminPin) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Admins');
  if (!sheet) {
    getOrCreateSheets();
    sheet = ss.getSheetByName('Admins');
  }

  var data = sheet.getDataRange().getValues();
  var id = adminData.id || ('admin_' + new Date().getTime());
  var role = adminData.role || 'sub_admin';
  var now = new Date().toISOString();

  var rowIndex = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(id).trim()) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex > -1) {
    if (adminData.name) sheet.getRange(rowIndex, 2).setValue(adminData.name);
    if (adminData.department) sheet.getRange(rowIndex, 3).setValue(adminData.department);
    if (adminData.contact) sheet.getRange(rowIndex, 4).setValue(adminData.contact);
    if (adminData.role) sheet.getRange(rowIndex, 5).setValue(adminData.role);
    if (adminData.pin) sheet.getRange(rowIndex, 6).setValue(String(adminData.pin));
    sheet.getRange(rowIndex, 8).setValue(now);
  } else {
    sheet.appendRow([
      id,
      adminData.name || 'แอดมินรอง',
      adminData.department || 'ผู้ดูแลระบบส่วนกลาง',
      adminData.contact || '',
      role,
      String(adminData.pin || '1234'),
      formatDateString(adminData.createdAt) || formatDateString(new Date()),
      now
    ]);
  }

  // หากเป็นแอดมินหลัก ให้อัปเดตรหัสใน Settings ด้วย
  if (role === 'primary' && adminData.pin) {
    updateSettingValue('admin_pin', String(adminData.pin));
  }

  SpreadsheetApp.flush();
  return { success: true, id: id };
}

function deleteAdmin(adminId, adminPin) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Admins');
  if (!sheet) return { success: false, message: 'ไม่พบชีต Admins' };

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(adminId).trim()) {
      if (String(data[i][4]) === 'primary') {
        return { success: false, message: 'ไม่สามารถลบบัญชีแอดมินหลักได้' };
      }
      sheet.deleteRow(i + 1);
      SpreadsheetApp.flush();
      return { success: true };
    }
  }
  return { success: false, message: 'ไม่พบบัญชีแอดมินที่ต้องการลบ' };
}

function changeAdminPin(adminId, newPin, currentPin) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Admins');
  if (!sheet) {
    getOrCreateSheets();
    sheet = ss.getSheetByName('Admins');
  }

  var data = sheet.getDataRange().getValues();
  var now = new Date().toISOString();
  var found = false;

  for (var i = 1; i < data.length; i++) {
    var rowId = String(data[i][0]).trim();
    if (rowId === String(adminId).trim() || (!adminId && String(data[i][4]) === 'primary')) {
      sheet.getRange(i + 1, 6).setValue(String(newPin));
      sheet.getRange(i + 1, 8).setValue(now);
      if (String(data[i][4]) === 'primary') {
        updateSettingValue('admin_pin', String(newPin));
      }
      found = true;
      break;
    }
  }

  if (!found) {
    // ถ้าไม่พบคอนเฟิร์ม ให้ปรับปรุงใน Settings
    updateSettingValue('admin_pin', String(newPin));
  }

  SpreadsheetApp.flush();
  return { success: true, message: 'เปลี่ยนรหัสผ่านเรียบร้อยแล้ว' };
}

// ------------------------------
// API: Events (กิจกรรม)
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
        createdBy: String(row[9] || 'ระบบ'),
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
  if (!sheet) {
    getOrCreateSheets();
    sheet = ss.getSheetByName('Events');
  }

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

  SpreadsheetApp.flush();
  return { success: true, id: id };
}

function deleteEvent(eventId, adminPin) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Events');
  if (!sheet) return { success: false, message: 'ไม่พบชีต Events' };

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(eventId)) {
      sheet.deleteRow(i + 1);
      SpreadsheetApp.flush();
      return { success: true };
    }
  }
  return { success: false, message: 'ไม่พบกิจกรรมที่ต้องการลบ' };
}

function batchSaveEvents(eventsList) {
  if (!Array.isArray(eventsList) || eventsList.length === 0) {
    return { success: false, message: 'ไม่มีรายการกิจกรรม' };
  }

  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Events');
  if (!sheet) {
    getOrCreateSheets();
    sheet = ss.getSheetByName('Events');
  }

  var data = sheet.getDataRange().getValues();
  var existingIds = {};
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) existingIds[String(data[i][0])] = i + 1;
  }

  var now = new Date().toISOString();
  var rowsToAppend = [];

  for (var k = 0; k < eventsList.length; k++) {
    var ev = eventsList[k];
    var id = ev.id || ('evt_' + new Date().getTime() + '_' + k);
    if (existingIds[id]) {
      var rIdx = existingIds[id];
      sheet.getRange(rIdx, 2).setValue(ev.title);
      sheet.getRange(rIdx, 3).setValue(ev.departmentId);
      sheet.getRange(rIdx, 4).setValue(ev.startDate);
      sheet.getRange(rIdx, 5).setValue(ev.endDate);
      sheet.getRange(rIdx, 6).setValue(ev.academicYear);
      sheet.getRange(rIdx, 7).setValue(ev.semester);
      sheet.getRange(rIdx, 8).setValue(ev.location || '');
      sheet.getRange(rIdx, 9).setValue(ev.description || '');
      sheet.getRange(rIdx, 11).setValue(ev.creatorName || '');
      sheet.getRange(rIdx, 12).setValue(ev.creatorDept || '');
      sheet.getRange(rIdx, 14).setValue(now);
    } else {
      rowsToAppend.push([
        id,
        ev.title,
        ev.departmentId,
        ev.startDate,
        ev.endDate,
        ev.academicYear,
        ev.semester,
        ev.location || '',
        ev.description || '',
        ev.createdBy || 'ระบบ',
        ev.creatorName || 'สำนักทะเบียนและประมวลผล',
        ev.creatorDept || 'ปฏิทิน ม.ฟาฏอนี',
        now,
        now
      ]);
    }
  }

  if (rowsToAppend.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
  }

  SpreadsheetApp.flush();
  return { success: true, count: eventsList.length };
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

function updateSettingValue(key, value) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Settings');
  if (!sheet) {
    getOrCreateSheets();
    sheet = ss.getSheetByName('Settings');
  }

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(key).trim()) {
      sheet.getRange(i + 1, 2).setValue(value);
      SpreadsheetApp.flush();
      return;
    }
  }
  sheet.appendRow([key, value]);
  SpreadsheetApp.flush();
}

function getSettings() {
  var ss = getSpreadsheet();
  var settings = { currentAcademicYear: '2568', adminPin: '1234' };
  var sheet = ss.getSheetByName('Settings');
  if (!sheet) return settings;

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === 'current_academic_year') {
      settings.currentAcademicYear = String(data[i][1]);
    } else if (data[i][0] === 'admin_pin') {
      settings.adminPin = String(data[i][1]);
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
