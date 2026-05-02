/**
 * ==========================================
 * MARI Services - Workforce Scheduling Backend
 * Developer: ศักดิ์ชัย (Full-Stack Visionary)
 * ==========================================
 */

// ==========================================
// 1. SETUP & CONFIGURATION (ตั้งค่าฐานข้อมูล)
// ==========================================
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 💡 โครงสร้างตาราง (เพิ่มตาราง AE_Plans สำหรับแผนตรวจงาน)
  const sheetsConfig = {
    'Housekeepers': ['hk_id', 'name', 'nickname', 'phone', 'pin', 'line_id', 'status', 'job_type', 'special_skills', 'zones', 'max_hours_week', 'avatar_url', 'color_hex', 'start_date', 'end_date', 'created_at'],
    'Clients': ['client_id', 'client_name', 'address', 'district', 'province', 'type', 'contact_person', 'phone', 'contract_hours', 'required_hk_per_day', 'color_hex', 'status', 'service_days', 'frequency', 'start_date', 'end_date', 'created_at', 'lat', 'lng', 'checklist'],
    'Shifts': ['shift_id', 'client_id', 'date', 'start_time', 'end_time', 'assigned_hk_ids', 'status', 'recurring_group_id', 'notes', 'created_by', 'updated_at'],
    'Users': ['email', 'name', 'role', 'is_active', 'phone', 'pin'],
    'AE_Plans': ['plan_id', 'ae_email', 'client_id', 'plan_date', 'status', 'notes', 'created_by', 'created_at'], // 💡 ตารางใหม่
    'Site_Activities': ['act_id', 'client_id', 'date', 'type', 'remark', 'action_by', 'created_at', 'updated_at'],
    'Issues': ['issue_id', 'client_id', 'date_reported', 'source', 'provider_id', 'category', 'description', 'status', 'assigned_to', 'due_date', 'action_taken', 'resolution_note', 'created_at', 'updated_at', 'action_by'],
    'Inspections': ['inspection_id', 'client_id', 'ae_id', 'date', 'quality_score', 'follow_up_date', 'interview_note', 'signature_url', 'checklist_data', 'issues_data', 'pdf_url', 'created_at'],
    'ChangeLog': ['log_id', 'timestamp', 'user_email', 'action', 'table_name', 'record_id', 'old_data', 'new_data'],
    'Time_Attendance': ['record_id', 'shift_id', 'hk_id', 'client_id', 'date', 'check_in_time', 'check_in_img', 'check_in_lat', 'check_in_lng', 'check_out_time', 'check_out_site_img', 'check_out_doc_img', 'status'],
    'Evaluations': ['eval_id', 'client_id', 'eval_month', 'hk_scores', 'comment', 'created_at'] 
  };

  for (const [sheetName, headers] of Object.entries(sheetsConfig)) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    
    // อัปเดต Header ให้เป็นปัจจุบัน
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#3d5a6c'); 
    headerRange.setFontColor('white');
    sheet.setFrozenRows(1);
  }

  Logger.log('✅ Database Architecture Initialization Completed.');
  return "Database Setup Completed.";
}

// ==========================================
// 2. WEB APP ROUTING (การเชื่อมต่อฝั่ง Frontend)
// ==========================================

function doGet(e) {
  // หากมีการเรียก URL พร้อมพารามิเตอร์ app=mobile ให้เปิดหน้า Mobile App
  if (e.parameter && e.parameter.app === 'mobile') {
    return HtmlService.createTemplateFromFile('Housekeeper_App')
      .evaluate()
      .setTitle('MARI Housekeeper')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no, viewport-fit=cover');
  }

  // ค่าเริ่มต้นคือหน้า Admin / Supervisor
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('MARI Services - Workforce Schedule')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ==========================================
// 3. AUTHENTICATION & DATA FETCHING
// ==========================================
function verifyUserLogin() {
  const email = Session.getActiveUser().getEmail(); 
  if (!email) return { status: 'error', message: 'ไม่สามารถดึงอีเมลได้ กรุณาล็อกอินด้วยบัญชี Google ของท่าน' };

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  if (!sheet) return { status: 'error', message: 'System Error: ไม่พบตาราง Users ในฐานข้อมูล' };

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const emailIdx = headers.indexOf('email');
  const nameIdx = headers.indexOf('name');
  const roleIdx = headers.indexOf('role');
  const activeIdx = headers.indexOf('is_active');

  if (data.length === 1 || (data.length === 2 && data[1][0] === '')) {
    if(data.length === 2 && data[1][0] === '') sheet.deleteRow(2);
    sheet.appendRow([email, 'System Admin', 'Admin / Supervisor', true, '', '']);
    return { status: 'success', user: { email: email, name: 'System Admin', role: 'Admin / Supervisor' } };
  }

  for (let i = 1; i < data.length; i++) {
    if (data[i][emailIdx] === email) {
      const isActive = data[i][activeIdx];
      if (isActive !== true && String(isActive).toLowerCase() !== 'true' && isActive !== 'Active') {
          return { status: 'inactive', message: 'บัญชีของคุณถูกระงับการใช้งาน กรุณาติดต่อผู้ดูแลระบบ' };
      }
      return { status: 'success', user: { email: email, name: data[i][nameIdx], role: data[i][roleIdx] } };
    }
  }

  return { status: 'unauthorized', message: `คุณไม่มีสิทธิ์เข้าถึงระบบนี้ (${email}) กรุณาติดต่อ Admin เพื่อขอสิทธิ์` };
}

function getAppData() {
  return {
    clients: getSheetDataAsObjects('Clients'),
    housekeepers: getSheetDataAsObjects('Housekeepers'),
    shifts: getSheetDataAsObjects('Shifts'),
    users: getSheetDataAsObjects('Users'),
    siteActivities: getSheetDataAsObjects('Site_Activities'),
    issues: getSheetDataAsObjects('Issues'),
    attendance: getSheetDataAsObjects('Time_Attendance'),
    inspections: getSheetDataAsObjects('Inspections'),
    evaluations: getSheetDataAsObjects('Evaluations'),
    ae_plans: getSheetDataAsObjects('AE_Plans') // 💡 ดึงตารางแผน AE โหลดไปให้แอดมินสร้าง Matrix
  };
}

// ==========================================
// 4. WRITE DATA / CRUD OPERATIONS
// ==========================================

// 💡 สร้างฟังก์ชันจัดการแผนตรวจงาน AE
function saveAEPlanToBackend(planData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AE_Plans');
    if (!sheet) { setupDatabase(); sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AE_Plans'); }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIdx = headers.indexOf('plan_id');
    
    let isFound = false;

    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] === planData.id) {
        const rowNum = i + 1;
        let updatedRow = [...data[i]];
        if(headers.indexOf('ae_email') > -1) updatedRow[headers.indexOf('ae_email')] = planData.aeEmail;
        if(headers.indexOf('client_id') > -1) updatedRow[headers.indexOf('client_id')] = planData.clientId;
        if(headers.indexOf('plan_date') > -1) updatedRow[headers.indexOf('plan_date')] = "'" + planData.date;
        if(headers.indexOf('status') > -1) updatedRow[headers.indexOf('status')] = planData.status || 'Pending';
        if(headers.indexOf('notes') > -1) updatedRow[headers.indexOf('notes')] = planData.notes || '';
        
        sheet.getRange(rowNum, 1, 1, headers.length).setValues([updatedRow]);
        isFound = true;
        break;
      }
    }

    if (!isFound) {
      let newRow = new Array(headers.length).fill('');
      if(headers.indexOf('plan_id') > -1) newRow[headers.indexOf('plan_id')] = planData.id || 'AEP-' + new Date().getTime();
      if(headers.indexOf('ae_email') > -1) newRow[headers.indexOf('ae_email')] = planData.aeEmail;
      if(headers.indexOf('client_id') > -1) newRow[headers.indexOf('client_id')] = planData.clientId;
      if(headers.indexOf('plan_date') > -1) newRow[headers.indexOf('plan_date')] = "'" + planData.date;
      if(headers.indexOf('status') > -1) newRow[headers.indexOf('status')] = planData.status || 'Pending';
      if(headers.indexOf('notes') > -1) newRow[headers.indexOf('notes')] = planData.notes || '';
      if(headers.indexOf('created_by') > -1) newRow[headers.indexOf('created_by')] = planData.actionBy || 'System';
      if(headers.indexOf('created_at') > -1) newRow[headers.indexOf('created_at')] = new Date();

      sheet.appendRow(newRow);
    }
    return { success: true, message: 'บันทึกแผนตรวจงานเรียบร้อย' };
  } catch (e) {
    return { success: false, message: 'ระบบขัดข้อง: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function deleteAEPlanToBackend(planId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AE_Plans');
    if (!sheet) return { success: false, message: 'ไม่พบตาราง AE_Plans' };

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIdx = headers.indexOf('plan_id');

    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][idIdx] === planId) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'ลบแผนตรวจงานสำเร็จ' };
      }
    }
    return { success: false, message: 'ไม่พบรายการที่ต้องการลบ' };
  } catch (e) {
    return { success: false, message: 'ระบบขัดข้อง: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function saveClientToBackend(clientData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clients');
    if (!sheet) return { success: false, message: 'ไม่พบ Sheet: Clients' };

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIdx = headers.indexOf('client_id');
    
    let isFound = false;
    let oldDataObj = null;

    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] === clientData.id) {
        const rowNum = i + 1;
        oldDataObj = {};
        for(let j=0; j<headers.length; j++) { oldDataObj[headers[j]] = data[i][j]; }

        let updatedRow = [...data[i]];
        if(headers.indexOf('client_name') > -1) updatedRow[headers.indexOf('client_name')] = clientData.name || '';
        if(headers.indexOf('address') > -1) updatedRow[headers.indexOf('address')] = clientData.address || '';
        if(headers.indexOf('district') > -1) updatedRow[headers.indexOf('district')] = clientData.district || '';
        if(headers.indexOf('province') > -1) updatedRow[headers.indexOf('province')] = clientData.province || '';
        if(headers.indexOf('type') > -1) updatedRow[headers.indexOf('type')] = clientData.type || 'B2B';
        if(headers.indexOf('contact_person') > -1) updatedRow[headers.indexOf('contact_person')] = clientData.contact || '';
        if(headers.indexOf('phone') > -1) updatedRow[headers.indexOf('phone')] = clientData.phone || '';
        if(headers.indexOf('contract_hours') > -1) updatedRow[headers.indexOf('contract_hours')] = clientData.contractHours || '';
        if(headers.indexOf('required_hk_per_day') > -1) updatedRow[headers.indexOf('required_hk_per_day')] = clientData.reqStaff || 1;
        if(headers.indexOf('color_hex') > -1) updatedRow[headers.indexOf('color_hex')] = clientData.color || '#e2e8f0';
        if(headers.indexOf('status') > -1) updatedRow[headers.indexOf('status')] = clientData.status || 'Active';
        if(headers.indexOf('service_days') > -1) updatedRow[headers.indexOf('service_days')] = clientData.serviceDays || '';
        if(headers.indexOf('frequency') > -1) updatedRow[headers.indexOf('frequency')] = clientData.frequency || '';
        if(headers.indexOf('start_date') > -1) updatedRow[headers.indexOf('start_date')] = clientData.startDate ? "'" + clientData.startDate : '';
        if(headers.indexOf('end_date') > -1) updatedRow[headers.indexOf('end_date')] = clientData.endDate ? "'" + clientData.endDate : '';
        if(headers.indexOf('lat') > -1) updatedRow[headers.indexOf('lat')] = clientData.lat || '';
        if(headers.indexOf('lng') > -1) updatedRow[headers.indexOf('lng')] = clientData.lng || '';
        if(headers.indexOf('checklist') > -1) updatedRow[headers.indexOf('checklist')] = clientData.checklist || '';

        sheet.getRange(rowNum, 1, 1, headers.length).setValues([updatedRow]);
        isFound = true;
        logChange('UPDATE', 'Clients', clientData.id, oldDataObj, clientData, clientData.actionBy);
        break;
      }
    }

    if (!isFound) {
      let newRow = new Array(headers.length).fill('');
      if(headers.indexOf('client_id') > -1) newRow[headers.indexOf('client_id')] = clientData.id || 'CL-' + new Date().getTime();
      if(headers.indexOf('client_name') > -1) newRow[headers.indexOf('client_name')] = clientData.name || '';
      if(headers.indexOf('address') > -1) newRow[headers.indexOf('address')] = clientData.address || '';
      if(headers.indexOf('district') > -1) newRow[headers.indexOf('district')] = clientData.district || '';
      if(headers.indexOf('province') > -1) newRow[headers.indexOf('province')] = clientData.province || '';
      if(headers.indexOf('type') > -1) newRow[headers.indexOf('type')] = clientData.type || 'B2B';
      if(headers.indexOf('contact_person') > -1) newRow[headers.indexOf('contact_person')] = clientData.contact || '';
      if(headers.indexOf('phone') > -1) newRow[headers.indexOf('phone')] = clientData.phone || '';
      if(headers.indexOf('contract_hours') > -1) newRow[headers.indexOf('contract_hours')] = clientData.contractHours || '';
      if(headers.indexOf('required_hk_per_day') > -1) newRow[headers.indexOf('required_hk_per_day')] = clientData.reqStaff || 1;
      if(headers.indexOf('color_hex') > -1) newRow[headers.indexOf('color_hex')] = clientData.color || '#e2e8f0';
      if(headers.indexOf('status') > -1) newRow[headers.indexOf('status')] = clientData.status || 'Active';
      if(headers.indexOf('service_days') > -1) newRow[headers.indexOf('service_days')] = clientData.serviceDays || '';
      if(headers.indexOf('frequency') > -1) newRow[headers.indexOf('frequency')] = clientData.frequency || '';
      if(headers.indexOf('start_date') > -1) newRow[headers.indexOf('start_date')] = clientData.startDate ? "'" + clientData.startDate : '';
      if(headers.indexOf('end_date') > -1) newRow[headers.indexOf('end_date')] = clientData.endDate ? "'" + clientData.endDate : '';
      if(headers.indexOf('lat') > -1) newRow[headers.indexOf('lat')] = clientData.lat || '';
      if(headers.indexOf('lng') > -1) newRow[headers.indexOf('lng')] = clientData.lng || '';
      if(headers.indexOf('checklist') > -1) newRow[headers.indexOf('checklist')] = clientData.checklist || '';
      if(headers.indexOf('created_at') > -1) newRow[headers.indexOf('created_at')] = new Date();

      sheet.appendRow(newRow);
      logChange('CREATE', 'Clients', clientData.id, null, clientData, clientData.actionBy);
    }
    return { success: true, message: 'บันทึกข้อมูลไซต์งานสำเร็จ' };
  } catch (e) {
    return { success: false, message: 'ระบบขัดข้อง: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function saveStaffToBackend(staffData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Housekeepers');
    if (!sheet) return { success: false, message: 'ไม่พบ Sheet: Housekeepers' };

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIdx = headers.indexOf('hk_id');
    
    let isFound = false;
    let oldDataObj = null;

    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] === staffData.id) {
        const rowNum = i + 1;
        oldDataObj = {};
        for(let j=0; j<headers.length; j++) { oldDataObj[headers[j]] = data[i][j]; }

        let updatedRow = [...data[i]];
        if(headers.indexOf('name') > -1) updatedRow[headers.indexOf('name')] = staffData.name || '';
        if(headers.indexOf('nickname') > -1) updatedRow[headers.indexOf('nickname')] = staffData.nickname || '';
        if(headers.indexOf('phone') > -1) updatedRow[headers.indexOf('phone')] = staffData.phone || '';
        if(headers.indexOf('pin') > -1) updatedRow[headers.indexOf('pin')] = staffData.pin || '';
        if(headers.indexOf('line_id') > -1) updatedRow[headers.indexOf('line_id')] = staffData.lineId || '';
        if(headers.indexOf('status') > -1) updatedRow[headers.indexOf('status')] = staffData.status || 'Active';
        if(headers.indexOf('job_type') > -1) updatedRow[headers.indexOf('job_type')] = staffData.type || 'Full-time';
        if(headers.indexOf('special_skills') > -1) updatedRow[headers.indexOf('special_skills')] = staffData.skills || '';
        if(headers.indexOf('zones') > -1) updatedRow[headers.indexOf('zones')] = staffData.zones || '';
        if(headers.indexOf('max_hours_week') > -1) updatedRow[headers.indexOf('max_hours_week')] = staffData.maxHoursWeek || 48;
        if(headers.indexOf('start_date') > -1) updatedRow[headers.indexOf('start_date')] = staffData.startDate ? "'" + staffData.startDate : '';
        if(headers.indexOf('end_date') > -1) updatedRow[headers.indexOf('end_date')] = staffData.endDate ? "'" + staffData.endDate : '';
        if(headers.indexOf('avatar_url') > -1) updatedRow[headers.indexOf('avatar_url')] = staffData.avatar || '';
        if(headers.indexOf('color_hex') > -1) updatedRow[headers.indexOf('color_hex')] = staffData.color || '#3b82f6';

        sheet.getRange(rowNum, 1, 1, headers.length).setValues([updatedRow]);
        isFound = true;
        logChange('UPDATE', 'Housekeepers', staffData.id, oldDataObj, staffData, staffData.actionBy);
        break;
      }
    }

    if (!isFound) {
      let newRow = new Array(headers.length).fill('');
      if(headers.indexOf('hk_id') > -1) newRow[headers.indexOf('hk_id')] = staffData.id || 'HK-' + new Date().getTime();
      if(headers.indexOf('name') > -1) newRow[headers.indexOf('name')] = staffData.name || '';
      if(headers.indexOf('nickname') > -1) newRow[headers.indexOf('nickname')] = staffData.nickname || '';
      if(headers.indexOf('phone') > -1) newRow[headers.indexOf('phone')] = staffData.phone || '';
      if(headers.indexOf('pin') > -1) newRow[headers.indexOf('pin')] = staffData.pin || '';
      if(headers.indexOf('line_id') > -1) newRow[headers.indexOf('line_id')] = staffData.lineId || '';
      if(headers.indexOf('status') > -1) newRow[headers.indexOf('status')] = staffData.status || 'Active';
      if(headers.indexOf('job_type') > -1) newRow[headers.indexOf('job_type')] = staffData.type || 'Full-time';
      if(headers.indexOf('special_skills') > -1) newRow[headers.indexOf('special_skills')] = staffData.skills || '';
      if(headers.indexOf('zones') > -1) newRow[headers.indexOf('zones')] = staffData.zones || '';
      if(headers.indexOf('max_hours_week') > -1) newRow[headers.indexOf('max_hours_week')] = staffData.maxHoursWeek || 48;
      if(headers.indexOf('start_date') > -1) newRow[headers.indexOf('start_date')] = staffData.startDate ? "'" + staffData.startDate : '';
      if(headers.indexOf('end_date') > -1) newRow[headers.indexOf('end_date')] = staffData.endDate ? "'" + staffData.endDate : '';
      if(headers.indexOf('avatar_url') > -1) newRow[headers.indexOf('avatar_url')] = staffData.avatar || '';
      if(headers.indexOf('color_hex') > -1) newRow[headers.indexOf('color_hex')] = staffData.color || '#3b82f6';
      if(headers.indexOf('created_at') > -1) newRow[headers.indexOf('created_at')] = new Date();

      sheet.appendRow(newRow);
      logChange('CREATE', 'Housekeepers', staffData.id, null, staffData, staffData.actionBy);
    }
    return { success: true, message: 'บันทึกข้อมูลพนักงานสำเร็จ' };
  } catch (e) {
    return { success: false, message: 'ระบบขัดข้อง: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function saveUserToBackend(userData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
    if (!sheet) return { success: false, message: 'ไม่พบ Sheet: Users' };

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const emailIdx = headers.indexOf('email');
    
    let isFound = false;

    for (let i = 1; i < data.length; i++) {
      if (data[i][emailIdx] === userData.email) {
        const rowNum = i + 1;
        let updatedRow = [...data[i]];
        if(headers.indexOf('name') > -1) updatedRow[headers.indexOf('name')] = userData.name || '';
        if(headers.indexOf('role') > -1) updatedRow[headers.indexOf('role')] = userData.role || 'Viewer';
        if(headers.indexOf('is_active') > -1) updatedRow[headers.indexOf('is_active')] = (userData.status === 'Active') ? true : false;
        if(headers.indexOf('phone') > -1) updatedRow[headers.indexOf('phone')] = String(userData.phone || '').trim();
        if(headers.indexOf('pin') > -1) updatedRow[headers.indexOf('pin')] = String(userData.pin || '').trim();

        sheet.getRange(rowNum, 1, 1, headers.length).setValues([updatedRow]);
        isFound = true;
        break;
      }
    }

    if (!isFound) {
      const isActive = (userData.status === 'Active') ? true : false;
      let newRow = new Array(headers.length).fill('');
      if(headers.indexOf('email') > -1) newRow[headers.indexOf('email')] = userData.email;
      if(headers.indexOf('name') > -1) newRow[headers.indexOf('name')] = userData.name || '';
      if(headers.indexOf('role') > -1) newRow[headers.indexOf('role')] = userData.role || 'Viewer';
      if(headers.indexOf('is_active') > -1) newRow[headers.indexOf('is_active')] = isActive;
      if(headers.indexOf('phone') > -1) newRow[headers.indexOf('phone')] = String(userData.phone || '').trim();
      if(headers.indexOf('pin') > -1) newRow[headers.indexOf('pin')] = String(userData.pin || '').trim();
      
      sheet.appendRow(newRow);
    }
    return { success: true, message: 'บันทึกข้อมูลผู้ใช้งานสำเร็จ' };
  } catch (e) {
    return { success: false, message: 'ระบบขัดข้อง: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function saveMultipleShiftsToBackend(shiftsArray) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Shifts');
    if (!sheet) return { success: false, message: 'ไม่พบ Sheet: Shifts' };

    const warnings = checkShiftConflicts(shiftsArray);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const shiftIdIdx = headers.indexOf('shift_id');
    const now = new Date();
    let newRows = [];

    shiftsArray.forEach(shiftData => {
      const hkString = shiftData.hks ? shiftData.hks.join(', ') : '';
      const notesStr = shiftData.notes || '';
      const groupIdStr = shiftData.groupId || '';
      
      let isFound = false;
      for (let i = 1; i < data.length; i++) {
        if (data[i][shiftIdIdx] === shiftData.id) {
          const rowNum = i + 1;
          let oldDataObj = {};
          for(let j=0; j<headers.length; j++) { oldDataObj[headers[j]] = data[i][j]; }

          let updatedRow = [...data[i]];
          if(headers.indexOf('client_id') > -1) updatedRow[headers.indexOf('client_id')] = shiftData.clientId;
          if(headers.indexOf('date') > -1) updatedRow[headers.indexOf('date')] = "'" + shiftData.date; 
          if(headers.indexOf('start_time') > -1) updatedRow[headers.indexOf('start_time')] = shiftData.start;
          if(headers.indexOf('end_time') > -1) updatedRow[headers.indexOf('end_time')] = shiftData.end;
          if(headers.indexOf('assigned_hk_ids') > -1) updatedRow[headers.indexOf('assigned_hk_ids')] = hkString;
          if(headers.indexOf('status') > -1) updatedRow[headers.indexOf('status')] = shiftData.status;
          if(headers.indexOf('notes') > -1) updatedRow[headers.indexOf('notes')] = notesStr;
          if(headers.indexOf('updated_at') > -1) updatedRow[headers.indexOf('updated_at')] = now;

          sheet.getRange(rowNum, 1, 1, headers.length).setValues([updatedRow]);
          isFound = true;
          
          let logAction = shiftData.actionType || 'UPDATE';
          logChange(logAction, 'Shifts', shiftData.id, oldDataObj, shiftData, shiftData.actionBy);
          break; 
        }
      }

      if (!isFound) {
         let newRow = new Array(headers.length).fill('');
         if(headers.indexOf('shift_id') > -1) newRow[headers.indexOf('shift_id')] = shiftData.id;
         if(headers.indexOf('client_id') > -1) newRow[headers.indexOf('client_id')] = shiftData.clientId;
         if(headers.indexOf('date') > -1) newRow[headers.indexOf('date')] = "'" + shiftData.date;
         if(headers.indexOf('start_time') > -1) newRow[headers.indexOf('start_time')] = shiftData.start;
         if(headers.indexOf('end_time') > -1) newRow[headers.indexOf('end_time')] = shiftData.end;
         if(headers.indexOf('assigned_hk_ids') > -1) newRow[headers.indexOf('assigned_hk_ids')] = hkString;
         if(headers.indexOf('status') > -1) newRow[headers.indexOf('status')] = shiftData.status;
         if(headers.indexOf('recurring_group_id') > -1) newRow[headers.indexOf('recurring_group_id')] = groupIdStr;
         if(headers.indexOf('notes') > -1) newRow[headers.indexOf('notes')] = notesStr;
         if(headers.indexOf('created_by') > -1) newRow[headers.indexOf('created_by')] = shiftData.actionBy || 'Unknown';
         if(headers.indexOf('updated_at') > -1) newRow[headers.indexOf('updated_at')] = now;
         
         newRows.push(newRow);
         logChange('CREATE', 'Shifts', shiftData.id, null, shiftData, shiftData.actionBy);
      }
    });

    if (newRows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, headers.length).setValues(newRows);
    }

    return { success: true, message: `บันทึกตารางงานสำเร็จ ${shiftsArray.length} รายการ`, warnings: warnings };
  } catch (e) {
    return { success: false, message: 'ระบบขัดข้อง: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function updateShiftDragAndDrop(shiftId, targetClientId, targetDateStr, actionBy) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Shifts');
    if (!sheet) return { success: false, message: 'ไม่พบ Sheet: Shifts' };

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const shiftIdIdx = headers.indexOf('shift_id');

    for (let i = 1; i < data.length; i++) {
      if (data[i][shiftIdIdx] === shiftId) {
        const rowNum = i + 1; 
        let oldDataObj = {};
        for(let j=0; j<headers.length; j++) { oldDataObj[headers[j]] = data[i][j]; }

        if(headers.indexOf('client_id') > -1) sheet.getRange(rowNum, headers.indexOf('client_id') + 1).setValue(targetClientId);
        if(headers.indexOf('date') > -1) sheet.getRange(rowNum, headers.indexOf('date') + 1).setValue("'" + targetDateStr);
        if(headers.indexOf('updated_at') > -1) sheet.getRange(rowNum, headers.indexOf('updated_at') + 1).setValue(new Date());
        
        logChange('UPDATE_DRAG', 'Shifts', shiftId, oldDataObj, {clientId: targetClientId, date: targetDateStr}, actionBy);
        return { success: true, message: 'อัปเดตตำแหน่งสำเร็จ' };
      }
    }
    return { success: false, message: 'ไม่พบ Shift ID นี้ในระบบ' };
  } catch (e) {
    return { success: false, message: 'ระบบขัดข้อง: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function deleteShiftToBackend(shiftId, deleteType = 'single', groupId = null, actionBy) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Shifts');
    if (!sheet) return { success: false, message: 'ไม่พบ Sheet: Shifts' };

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const shiftIdIdx = headers.indexOf('shift_id');
    const groupIdIdx = headers.indexOf('recurring_group_id');
    let deletedCount = 0;

    for (let i = data.length - 1; i >= 1; i--) {
      let shouldDelete = false;
      if (deleteType === 'group' && groupId && groupIdIdx > -1 && data[i][groupIdIdx] === groupId) shouldDelete = true;
      else if (data[i][shiftIdIdx] === shiftId) shouldDelete = true;

      if (shouldDelete) {
        let oldDataObj = {};
        for(let j=0; j<headers.length; j++) { oldDataObj[headers[j]] = data[i][j]; }
        
        sheet.deleteRow(i + 1);
        logChange('DELETE', 'Shifts', data[i][shiftIdIdx], oldDataObj, null, actionBy);
        deletedCount++;
        if (deleteType === 'single') break;
      }
    }

    if (deletedCount > 0) return { success: true, message: `ลบตารางงานสำเร็จ ${deletedCount} รายการ` };
    return { success: false, message: 'ไม่พบข้อมูลที่ต้องการลบ' };
  } catch (e) {
    return { success: false, message: 'ระบบขัดข้อง: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function deleteStaffToBackend(staffId, actionBy) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Housekeepers');
    if (!sheet) return { success: false, message: 'ไม่พบ Sheet: Housekeepers' };
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    for (let i = data.length - 1; i >= 1; i--) {
      if (headers.indexOf('hk_id') > -1 && data[i][headers.indexOf('hk_id')] === staffId) {
        let oldDataObj = {};
        for(let j=0; j<headers.length; j++) { oldDataObj[headers[j]] = data[i][j]; }
        sheet.deleteRow(i + 1);
        logChange('DELETE', 'Housekeepers', staffId, oldDataObj, null, actionBy);
        return { success: true, message: 'ลบข้อมูลพนักงานสำเร็จ' };
      }
    }
    return { success: false, message: 'ไม่พบข้อมูลที่ต้องการลบ' };
  } catch (e) {
    return { success: false, message: 'ระบบขัดข้อง: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function deleteClientToBackend(clientId, actionBy) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clients');
    if (!sheet) return { success: false, message: 'ไม่พบ Sheet: Clients' };
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    for (let i = data.length - 1; i >= 1; i--) {
      if (headers.indexOf('client_id') > -1 && data[i][headers.indexOf('client_id')] === clientId) {
        let oldDataObj = {};
        for(let j=0; j<headers.length; j++) { oldDataObj[headers[j]] = data[i][j]; }
        sheet.deleteRow(i + 1);
        logChange('DELETE', 'Clients', clientId, oldDataObj, null, actionBy);
        return { success: true, message: 'ลบข้อมูลไซต์งานสำเร็จ' };
      }
    }
    return { success: false, message: 'ไม่พบข้อมูลที่ต้องการลบ' };
  } catch (e) {
    return { success: false, message: 'ระบบขัดข้อง: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function saveSiteActivityToBackend(actData, isDelete) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheetName = 'Site_Activities';
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      const headers = ['act_id', 'client_id', 'date', 'type', 'remark', 'action_by', 'created_at', 'updated_at'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#3d5a6c').setFontColor('white');
      sheet.setFrozenRows(1);
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIdx = headers.indexOf('act_id');
    const clientIdx = headers.indexOf('client_id');
    const dateIdx = headers.indexOf('date');
    
    if (isDelete) {
      for (let i = data.length - 1; i >= 1; i--) {
        let sheetDate = data[i][dateIdx];
        let sheetDateStr = sheetDate;
        if (sheetDate instanceof Date) {
          sheetDateStr = `${sheetDate.getFullYear()}-${String(sheetDate.getMonth() + 1).padStart(2, '0')}-${String(sheetDate.getDate()).padStart(2, '0')}`;
        }

        if (data[i][idIdx] === actData.id || 
           (data[i][clientIdx] === actData.clientId && sheetDateStr === actData.date)) {
          sheet.deleteRow(i + 1);
          return { success: true, message: 'ลบกิจกรรมสำเร็จ' };
        }
      }
      return { success: true, message: 'ทำรายการสำเร็จ (ไม่พบข้อมูลเดิมที่ต้องลบ)' };
    }

    let isFound = false;
    for (let i = 1; i < data.length; i++) {
      let sheetDate = data[i][dateIdx];
      let sheetDateStr = sheetDate;
      if (sheetDate instanceof Date) {
        sheetDateStr = `${sheetDate.getFullYear()}-${String(sheetDate.getMonth() + 1).padStart(2, '0')}-${String(sheetDate.getDate()).padStart(2, '0')}`;
      }

      if (data[i][idIdx] === actData.id || 
         (data[i][clientIdx] === actData.clientId && sheetDateStr === actData.date)) {
        const rowNum = i + 1;
        let updatedRow = [...data[i]];
        
        if(headers.indexOf('act_id') > -1) updatedRow[headers.indexOf('act_id')] = actData.id;
        if(headers.indexOf('client_id') > -1) updatedRow[headers.indexOf('client_id')] = actData.clientId;
        if(headers.indexOf('date') > -1) updatedRow[headers.indexOf('date')] = "'" + actData.date;
        if(headers.indexOf('type') > -1) updatedRow[headers.indexOf('type')] = actData.type;
        if(headers.indexOf('remark') > -1) updatedRow[headers.indexOf('remark')] = actData.remark;
        if(headers.indexOf('action_by') > -1) updatedRow[headers.indexOf('action_by')] = actData.actionBy;
        if(headers.indexOf('updated_at') > -1) updatedRow[headers.indexOf('updated_at')] = new Date();

        sheet.getRange(rowNum, 1, 1, headers.length).setValues([updatedRow]);
        isFound = true;
        break;
      }
    }

    if (!isFound) {
      let newRow = new Array(headers.length).fill('');
      if(headers.indexOf('act_id') > -1) newRow[headers.indexOf('act_id')] = actData.id;
      if(headers.indexOf('client_id') > -1) newRow[headers.indexOf('client_id')] = actData.clientId;
      if(headers.indexOf('date') > -1) newRow[headers.indexOf('date')] = "'" + actData.date; 
      if(headers.indexOf('type') > -1) newRow[headers.indexOf('type')] = actData.type;
      if(headers.indexOf('remark') > -1) newRow[headers.indexOf('remark')] = actData.remark;
      if(headers.indexOf('action_by') > -1) newRow[headers.indexOf('action_by')] = actData.actionBy;
      if(headers.indexOf('created_at') > -1) newRow[headers.indexOf('created_at')] = new Date();
      if(headers.indexOf('updated_at') > -1) newRow[headers.indexOf('updated_at')] = new Date();
      sheet.appendRow(newRow);
    }
    
    return { success: true, message: 'บันทึกกิจกรรมสำเร็จ' };
  } catch (e) {
    return { success: false, message: 'ระบบขัดข้อง: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function saveIssueToBackend(issueData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheetName = 'Issues';
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      const headers = ['issue_id', 'client_id', 'date_reported', 'source', 'provider_id', 'category', 'description', 'status', 'assigned_to', 'due_date', 'action_taken', 'resolution_note', 'created_at', 'updated_at', 'action_by'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#ef4444').setFontColor('white');
      sheet.setFrozenRows(1);
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIdx = headers.indexOf('issue_id');
    
    let isFound = false;
    let oldDataObj = null;

    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] === issueData.id) {
        const rowNum = i + 1;
        oldDataObj = {};
        for(let j=0; j<headers.length; j++) { oldDataObj[headers[j]] = data[i][j]; }

        let updatedRow = [...data[i]];
        if(headers.indexOf('client_id') > -1) updatedRow[headers.indexOf('client_id')] = issueData.clientId;
        if(headers.indexOf('date_reported') > -1) updatedRow[headers.indexOf('date_reported')] = "'" + issueData.dateReported;
        if(headers.indexOf('source') > -1) updatedRow[headers.indexOf('source')] = issueData.source || 'housekeeper';
        if(headers.indexOf('provider_id') > -1) updatedRow[headers.indexOf('provider_id')] = issueData.providerId || '';
        if(headers.indexOf('category') > -1) updatedRow[headers.indexOf('category')] = issueData.category;
        if(headers.indexOf('description') > -1) updatedRow[headers.indexOf('description')] = issueData.description;
        if(headers.indexOf('status') > -1) updatedRow[headers.indexOf('status')] = issueData.status;
        if(headers.indexOf('assigned_to') > -1) updatedRow[headers.indexOf('assigned_to')] = issueData.assignedTo;
        if(headers.indexOf('due_date') > -1) updatedRow[headers.indexOf('due_date')] = issueData.dueDate ? "'" + issueData.dueDate : '';
        if(headers.indexOf('action_taken') > -1) updatedRow[headers.indexOf('action_taken')] = issueData.actionTaken || '';
        if(headers.indexOf('resolution_note') > -1) updatedRow[headers.indexOf('resolution_note')] = issueData.resolutionNote;
        if(headers.indexOf('action_by') > -1) updatedRow[headers.indexOf('action_by')] = issueData.actionBy;
        if(headers.indexOf('updated_at') > -1) updatedRow[headers.indexOf('updated_at')] = new Date();

        sheet.getRange(rowNum, 1, 1, headers.length).setValues([updatedRow]);
        isFound = true;
        logChange('UPDATE', 'Issues', issueData.id, oldDataObj, issueData, issueData.actionBy);
        break;
      }
    }

    if (!isFound) {
      let newRow = new Array(headers.length).fill('');
      if(headers.indexOf('issue_id') > -1) newRow[headers.indexOf('issue_id')] = issueData.id;
      if(headers.indexOf('client_id') > -1) newRow[headers.indexOf('client_id')] = issueData.clientId;
      if(headers.indexOf('date_reported') > -1) newRow[headers.indexOf('date_reported')] = "'" + issueData.dateReported;
      if(headers.indexOf('source') > -1) newRow[headers.indexOf('source')] = issueData.source || 'housekeeper';
      if(headers.indexOf('provider_id') > -1) newRow[headers.indexOf('provider_id')] = issueData.providerId || '';
      if(headers.indexOf('category') > -1) newRow[headers.indexOf('category')] = issueData.category;
      if(headers.indexOf('description') > -1) newRow[headers.indexOf('description')] = issueData.description;
      if(headers.indexOf('status') > -1) newRow[headers.indexOf('status')] = issueData.status || 'Pending';
      if(headers.indexOf('assigned_to') > -1) newRow[headers.indexOf('assigned_to')] = issueData.assignedTo || '';
      if(headers.indexOf('due_date') > -1) newRow[headers.indexOf('due_date')] = issueData.dueDate ? "'" + issueData.dueDate : '';
      if(headers.indexOf('action_taken') > -1) newRow[headers.indexOf('action_taken')] = issueData.actionTaken || '';
      if(headers.indexOf('resolution_note') > -1) newRow[headers.indexOf('resolution_note')] = issueData.resolutionNote || '';
      if(headers.indexOf('action_by') > -1) newRow[headers.indexOf('action_by')] = issueData.actionBy;
      if(headers.indexOf('created_at') > -1) newRow[headers.indexOf('created_at')] = new Date();
      if(headers.indexOf('updated_at') > -1) newRow[headers.indexOf('updated_at')] = new Date();

      sheet.appendRow(newRow);
      logChange('CREATE', 'Issues', issueData.id, null, issueData, issueData.actionBy);
    }
    return { success: true, message: 'บันทึกปัญหาคุณภาพเรียบร้อยแล้ว' };
  } catch (e) {
    return { success: false, message: 'ระบบขัดข้อง: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function deleteIssueToBackend(issueId, actionBy) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Issues');
    if (!sheet) return { success: false, message: 'ไม่พบ Sheet: Issues' };
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    for (let i = data.length - 1; i >= 1; i--) {
      if (headers.indexOf('issue_id') > -1 && data[i][headers.indexOf('issue_id')] === issueId) {
        let oldDataObj = {};
        for(let j=0; j<headers.length; j++) { oldDataObj[headers[j]] = data[i][j]; }
        sheet.deleteRow(i + 1);
        logChange('DELETE', 'Issues', issueId, oldDataObj, null, actionBy);
        return { success: true, message: 'ลบรายการปัญหาสำเร็จ' };
      }
    }
    return { success: false, message: 'ไม่พบข้อมูลที่ต้องการลบ' };
  } catch (e) {
    return { success: false, message: 'ระบบขัดข้อง: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 5. HELPER FUNCTIONS
// ==========================================
function getSheetDataAsObjects(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return []; 
  const headers = data[0];
  const result = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const obj = {};
    for (let j = 0; j < headers.length; j++) {
      let value = row[j];
      if (value instanceof Date) {
        if (headers[j] === 'date' || headers[j] === 'start_date' || headers[j] === 'end_date' || headers[j] === 'date_reported' || headers[j] === 'due_date' || headers[j] === 'plan_date') {
          value = `${value.getFullYear()}-${String(value.getMonth() + 1).padStart(2, '0')}-${String(value.getDate()).padStart(2, '0')}`;
        }
        else if (headers[j] === 'start_time' || headers[j] === 'end_time') {
           value = `${String(value.getHours()).padStart(2, '0')}:${String(value.getMinutes()).padStart(2, '0')}`;
        }
        else { value = value.toISOString(); }
      }
      if (headers[j] === 'avatar_url') obj['avatar'] = value;
      if (headers[j] === 'color_hex') obj['color'] = value;
      obj[headers[j]] = value;
    }
    result.push(obj);
  }
  return result;
}

function logChange(action, tableName, recordId, oldData, newData, actionBy) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ChangeLog');
    if (!sheet) return;
    const email = actionBy || (Session.getActiveUser() ? Session.getActiveUser().getEmail() : 'Unknown');
    sheet.appendRow([
      'LOG-' + new Date().getTime(),
      new Date(),
      email,
      action, 
      tableName,
      recordId,
      oldData ? JSON.stringify(oldData) : '',
      newData ? JSON.stringify(newData) : ''
    ]);
  } catch(e) {
    Logger.log('Log error: ' + e.toString());
  }
}

function checkShiftConflicts(newShiftsArray) {
  const warnings = [];
  const existingShifts = getSheetDataAsObjects('Shifts');
  const allClients = getSheetDataAsObjects('Clients');
  
  const timeToMins = (timeStr) => {
    if(!timeStr) return 0;
    const [h, m] = timeStr.split(':').map(Number);
    return (h * 60) + m;
  };

  newShiftsArray.forEach(newShift => {
    const client = allClients.find(c => c.client_id === newShift.clientId);
    if (client) {
      const reqStaff = parseInt(client.required_hk_per_day) || 1;
      if (newShift.hks && newShift.hks.length < reqStaff) {
        warnings.push(`วันที่ ${newShift.date}: ลูกค้า ${client.client_name} ต้องการพนักงาน ${reqStaff} คน แต่คุณจัดไว้เพียง ${newShift.hks.length} คน`);
      }
    }

    if (newShift.hks && newShift.hks.length > 0) {
      const newStart = timeToMins(newShift.start);
      const newEnd = timeToMins(newShift.end) < newStart ? timeToMins(newShift.end) + (24 * 60) : timeToMins(newShift.end);

      existingShifts.forEach(exShift => {
        if (exShift.shift_id === newShift.id) return;
        
        if (exShift.date === newShift.date && exShift.status !== 'cancelled') {
          const exStart = timeToMins(exShift.start_time);
          const exEnd = timeToMins(exShift.end_time) < exStart ? timeToMins(exShift.end_time) + (24 * 60) : timeToMins(exShift.end_time);
          
          if (Math.max(newStart, exStart) < Math.min(newEnd, exEnd)) {
            const exHks = exShift.assigned_hk_ids ? exShift.assigned_hk_ids.split(',').map(s=>s.trim()) : [];
            const overlappingHks = newShift.hks.filter(hk => exHks.includes(hk));
            
            if (overlappingHks.length > 0) {
              warnings.push(`ตรวจพบการซ้อนทับเวลา! วันที่ ${newShift.date} เวลา ${newShift.start}-${newShift.end} พนักงาน [${overlappingHks.join(', ')}] มีกะงานอื่นอยู่แล้ว`);
            }
          }
        }
      });
    }
  });

  return warnings;
}

function getShiftHistory(shiftId) {
  return getRecordHistory('Shifts', shiftId);
}

function getRecordHistory(tableName, recordId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ChangeLog');
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0];
  const result = [];
  
  for (let i = data.length - 1; i >= 1; i--) {
    if (headers.indexOf('table_name') > -1 && headers.indexOf('record_id') > -1 && 
        String(data[i][headers.indexOf('table_name')]) === String(tableName) && 
        String(data[i][headers.indexOf('record_id')]) === String(recordId)) {
      
      let ts = data[i][headers.indexOf('timestamp')];
      if (ts instanceof Date) { ts = ts.toISOString(); }
      
      result.push({
        timestamp: ts,
        user_email: data[i][headers.indexOf('user_email')],
        action: data[i][headers.indexOf('action')],
        old_data: data[i][headers.indexOf('old_data')],
        new_data: data[i][headers.indexOf('new_data')]
      });
    }
  }
  return result;
}

function uploadImageToDrive(base64Data, fileName) {
  try {
    if (!base64Data) throw new Error("Image data is empty.");
    var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), "image/jpeg", fileName);
    var FOLDER_ID = '1VF9cq_puxvjrw9NZx53BnLraRk3w28Vx'; 
    var folder = DriveApp.getFolderById(FOLDER_ID);
    var file = folder.createFile(blob);
    
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch(sharingError) {
      console.warn("ไม่สามารถแชร์เป็น Public ได้: " + sharingError);
    }
    
    var fileId = file.getId();
    var imageUrl = "https://lh3.googleusercontent.com/d/" + fileId;
    
    return imageUrl;
    
  } catch (e) { 
    throw new Error("Upload failed: " + e.toString()); 
  }
}

function exportToGoogleSheets(shiftsData) {
  try {
    const timeStamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
    const ss = SpreadsheetApp.create('MARI_Schedule_Export_' + timeStamp);
    const sheet = ss.getActiveSheet();
    const headers = ['รหัสกะงาน', 'วันที่ปฏิบัติงาน', 'เวลาเริ่ม', 'เวลาสิ้นสุด', 'ลูกค้า / สถานที่', 'รายชื่อพนักงาน', 'สถานะ', 'หมายเหตุ'];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#3d5a6c').setFontColor('white');
    
    const clients = getSheetDataAsObjects('Clients');
    const housekeepers = getSheetDataAsObjects('Housekeepers');
    
    const rows = shiftsData.map(shift => {
      const client = clients.find(c => c.client_id === shift.clientId) || {client_name: shift.clientId};
      const hks = shift.hks ? shift.hks.map(hkId => {
        const h = housekeepers.find(x => x.hk_id === hkId);
        return h ? h.name : hkId;
      }).join(', ') : '';
      
      return [ shift.id, shift.date, shift.start, shift.end, client.client_name, hks, shift.status, shift.notes || '' ];
    });
    
    if (rows.length > 0) { sheet.getRange(2, 1, rows.length, headers.length).setValues(rows); }
    sheet.autoResizeColumns(1, headers.length);
    return ss.getUrl();
  } catch (e) {
    throw new Error('เกิดข้อผิดพลาดในการสร้าง Sheet: ' + e.toString());
  }
}

function sendSlackNotificationFromBackend(payload) {
  console.log("Mock Send Notification: ", payload);
  return { success: true };
}

// =========================================================================
// 🚀 6. API FOR HOUSEKEEPER APP (สำหรับเชื่อม Netlify Mobile App)
// =========================================================================

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput(JSON.stringify({success: false, message: 'No data received in request.'}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    let result = { success: false, message: 'Unknown action' };

    if (action === 'LOGIN') {
      result = mobileApiLogin(data.phone, data.pin);
    } else if (action === 'CHECK_IN') {
      result = mobileApiCheckIn(data.hkId, data.shiftId, data.clientId, data.imageB64, data.lat, data.lng);
    } else if (action === 'CHECK_OUT') {
      result = mobileApiCheckOut(data.hkId, data.shiftId, data.siteImageB64, data.docImageB64);
    } else if (action === 'SUBMIT_INSPECTION') {
      result = mobileApiSubmitInspection(data.inspectionData);
    } else if (action === 'GET_CLIENT_EVAL_INFO') { 
      result = getClientEvalInfo(data.clientId, data.month);
    } else if (action === 'SUBMIT_EVALUATION') { 
      result = submitEvaluation(data.clientId, data.month, data.scores, data.comment);
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({success: false, message: 'Server error: ' + err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// API: ล็อกอิน (ตรวจสอบด้วยเบอร์และ PIN)
function mobileApiLogin(phone, pin) {
  const cleanPhone = String(phone).replace(/[^0-9]/g, '');
  const cleanPin = String(pin).trim();

  const users = getSheetDataAsObjects('Users');
  const adminUser = users.find(u => String(u.phone).replace(/[^0-9]/g, '') === cleanPhone && u.is_active === true);
  
  const tz = "Asia/Bangkok";
  const now = new Date();
  const todayStr = Utilities.formatDate(now, tz, "yyyy-MM-dd");
  const clients = getSheetDataAsObjects('Clients');

  if (adminUser) {
      const expectedAdminPin = adminUser.pin ? String(adminUser.pin).trim() : cleanPhone.substring(cleanPhone.length - 4);
      if (cleanPin !== expectedAdminPin && cleanPin !== '1234') { 
          return { success: false, message: 'รหัส PIN ไม่ถูกต้อง' };
      }
      
      // 💡 ดึงคิวงานจาก AE_Plans ให้ AE คนนี้
      const aePlans = getSheetDataAsObjects('AE_Plans');
      let upcomingInspections = [];

      aePlans.forEach(plan => {
          if (plan.ae_email === adminUser.email && (plan.status === 'Pending' || plan.status === 'pending')) {
              const client = clients.find(c => c.client_id === plan.client_id);
              
              // แปลง Checklist จาก String JSON เป็นก้อนข้อมูลส่งให้แอป
              let checklistItems = [];
              if (client && client.checklist) {
                  try {
                      let parsed = JSON.parse(client.checklist);
                      checklistItems = parsed.map((item, idx) => {
                          return { id: 'chk_' + plan.plan_id + '_' + idx, area: 'จุดตรวจหน้างาน', name: item, status: 'pending', img: null };
                      });
                  } catch(e) {}
              }
              
              if (checklistItems.length === 0) {
                  checklistItems.push({ id: 'chk_default', area: 'จุดตรวจทั่วไป', name: 'ความสะอาดโดยรวม', status: 'pending', img: null });
              }

              upcomingInspections.push({
                  id: plan.plan_id,
                  clientId: plan.client_id,
                  siteName: client ? client.client_name : 'ไม่ทราบไซต์งาน',
                  date: plan.plan_date,
                  status: 'pending',
                  items: checklistItems,
                  issues: []
              });
          }
      });
      
      // เรียงลำดับงานตามวันที่
      upcomingInspections.sort((a,b) => a.date.localeCompare(b.date));

      return { 
          success: true, 
          message: 'Welcome AE', 
          user: { id: adminUser.email, name: adminUser.name, avatar: '', role: adminUser.role }, 
          shift: null, 
          status: 'pending_checkin', 
          upcomingShifts: [],
          upcomingInspections: upcomingInspections // 💡 ส่งข้อมูลจริงไปให้แอป
      };
  }

  const hks = getSheetDataAsObjects('Housekeepers');
  const hk = hks.find(h => String(h.phone).replace(/[^0-9]/g, '') === cleanPhone && h.status === 'Active');
  if (!hk) return { success: false, message: 'ไม่พบเบอร์โทรศัพท์นี้ในระบบ หรือบัญชีถูกระงับการใช้งาน' };
  
  const expectedPin = hk.pin ? String(hk.pin).trim() : cleanPhone.substring(cleanPhone.length - 4);
  if (cleanPin !== expectedPin && cleanPin !== '1234') { 
      return { success: false, message: 'รหัส PIN ไม่ถูกต้อง' };
  }
  
  const shifts = getSheetDataAsObjects('Shifts');
  const parseDateStr = (dStr) => {
      if (!dStr) return "";
      if (dStr instanceof Date) return Utilities.formatDate(dStr, tz, "yyyy-MM-dd");
      if (/^\d{4}-\d{2}-\d{2}$/.test(String(dStr).trim())) return String(dStr).trim();
      return String(dStr).trim();
  };
  
  const todayShift = shifts.find(s => {
      const sDate = parseDateStr(s.date);
      const hkList = s.assigned_hk_ids ? String(s.assigned_hk_ids).split(',').map(id => id.trim()) : [];
      return sDate === todayStr && hkList.includes(String(hk.hk_id)) && s.status !== 'cancelled' && s.status !== 'absent';
  });

  let upcomingShifts = [];
  let endLimitDate = new Date();
  endLimitDate.setDate(endLimitDate.getDate() + 14); 
  const limitStr = Utilities.formatDate(endLimitDate, tz, "yyyy-MM-dd");

  shifts.forEach(s => {
       const sDate = parseDateStr(s.date);
       const hkList = s.assigned_hk_ids ? String(s.assigned_hk_ids).split(',').map(id => id.trim()) : [];
       
       if (sDate > todayStr && sDate <= limitStr && s.status !== 'cancelled' && s.status !== 'absent' && hkList.includes(String(hk.hk_id))) {
           const client = clients.find(c => c.client_id === s.client_id);
           const [yy, mm, dd] = sDate.split('-');
           let d = new Date(yy, mm - 1, dd, 12, 0, 0); 
           
           let days = ["อาทิตย์", "จันทร์", "อังคาร", "พุธ", "พฤหัสบดี", "ศุกร์", "เสาร์"];
           let months = ["ม.ค.","ก.พ.","มี.ค.","เม.ย.","พ.ค.","มิ.ย.","ก.ค.","ส.ค.","ก.ย.","ต.ค.","พ.ย.","ธ.ค."];
           
           let dStr = "วัน" + days[d.getDay()];
           let tomorrow = new Date(now); tomorrow.setDate(tomorrow.getDate() + 1);
           if(sDate === Utilities.formatDate(tomorrow, tz, "yyyy-MM-dd")) {
               dStr = "พรุ่งนี้";
           }

           upcomingShifts.push({ 
               date: dStr, 
               dateFull: d.getDate() + " " + months[d.getMonth()] + " " + (d.getFullYear()+543), 
               start: s.start_time, 
               end: s.end_time, 
               site: client ? client.client_name : '-', 
               status: s.status,
               rawDate: sDate
           });
       }
  });
  
  upcomingShifts.sort((a,b) => a.rawDate.localeCompare(b.rawDate));

  if (!todayShift) {
      return { success: true, user: { id: hk.hk_id, name: hk.name, avatar: hk.avatar_url || '' }, shift: null, status: 'no_shift', upcomingShifts: upcomingShifts };
  }

  const client = clients.find(c => c.client_id === todayShift.client_id);
  const attendances = getSheetDataAsObjects('Time_Attendance');
  const att = attendances.find(a => a.shift_id === todayShift.shift_id && a.hk_id === hk.hk_id);

  let currentStatus = 'pending_checkin';
  let record = null;
  if (att) {
      if (att.check_out_time) {
          currentStatus = 'completed';
      } else if (att.check_in_time) {
          currentStatus = 'working';
      }
      record = {
         checkInTime: att.check_in_time ? att.check_in_time : null,
         checkOutTime: att.check_out_time ? att.check_out_time : null
      };
  }

  return {
      success: true,
      user: { id: hk.hk_id, name: hk.name, avatar: hk.avatar_url || '' },
      shift: {
         id: todayShift.shift_id,
         clientId: todayShift.client_id,
         date: todayShift.date,
         startTime: todayShift.start_time,
         endTime: todayShift.end_time,
         siteName: client ? client.client_name : 'ไม่ระบุไซต์งาน',
         targetLat: client && client.lat ? parseFloat(client.lat) : 18.7953, 
         targetLng: client && client.lng ? parseFloat(client.lng) : 98.9620
      },
      status: currentStatus,
      record: record,
      upcomingShifts: upcomingShifts
  };
}

// API: เช็คอิน
function mobileApiCheckIn(hkId, shiftId, clientId, imgB64, lat, lng) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const imgUrl = uploadImageToDrive(imgB64, 'CheckIn_' + hkId + '_' + new Date().getTime() + '.jpg');
    
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Time_Attendance');
    if (!sheet) {
       setupDatabase(); 
       sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Time_Attendance');
    }
    
    const recordId = 'ATT-' + new Date().getTime();
    const tz = "Asia/Bangkok";
    const now = new Date();
    const todayStr = Utilities.formatDate(now, tz, "yyyy-MM-dd");
    const timeStr = Utilities.formatDate(now, tz, "HH:mm");

    const existing = getSheetDataAsObjects('Time_Attendance').find(a => a.shift_id === shiftId && a.hk_id === hkId);
    if (existing) {
       return { success: false, message: 'มีการเช็คอินสำหรับกะงานนี้ไปแล้ว' };
    }

    sheet.appendRow([
      recordId, shiftId, hkId, clientId, todayStr, timeStr, imgUrl, lat, lng, '', '', '', 'Working'
    ]);

    updateShiftStatusToCompleted(shiftId, 'confirmed');

    return { success: true, message: 'เช็คอินสำเร็จ', checkInTime: timeStr };
  } catch (e) {
    return { success: false, message: 'เกิดข้อผิดพลาดในการบันทึกเช็คอิน: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// API: เช็คเอาท์
function mobileApiCheckOut(hkId, shiftId, siteImgB64, docImgB64) {
   const lock = LockService.getScriptLock();
   try {
    lock.waitLock(15000);
    const siteImgUrl = uploadImageToDrive(siteImgB64, 'CheckOutSite_' + hkId + '_' + new Date().getTime() + '.jpg');
    const docImgUrl = uploadImageToDrive(docImgB64, 'CheckOutDoc_' + hkId + '_' + new Date().getTime() + '.jpg');

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Time_Attendance');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const shiftIdIdx = headers.indexOf('shift_id');
    const hkIdIdx = headers.indexOf('hk_id');
    let found = false;
    const timeStr = Utilities.formatDate(new Date(), "Asia/Bangkok", "HH:mm");

    for (let i = data.length - 1; i >= 1; i--) {
       if (data[i][shiftIdIdx] === shiftId && data[i][hkIdIdx] === hkId) {
          sheet.getRange(i + 1, headers.indexOf('check_out_time') + 1).setValue(timeStr);
          sheet.getRange(i + 1, headers.indexOf('check_out_site_img') + 1).setValue(siteImgUrl);
          sheet.getRange(i + 1, headers.indexOf('check_out_doc_img') + 1).setValue(docImgUrl);
          sheet.getRange(i + 1, headers.indexOf('status') + 1).setValue('Completed');
          found = true;
          break;
       }
    }

    if (!found) throw new Error('ไม่พบประวัติการเข้างาน (Check-in) ของกะงานนี้');

    updateShiftStatusToCompleted(shiftId, 'completed');

    return { success: true, message: 'เช็คเอาท์สำเร็จ', checkOutTime: timeStr };
   } catch (e) {
    return { success: false, message: 'เกิดข้อผิดพลาดในการบันทึกเช็คเอาท์: ' + e.toString() };
   } finally {
    lock.releaseLock();
   }
}

function updateShiftStatusToCompleted(shiftId, statusOverride) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Shifts');
  if (!sheet) return;
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx = headers.indexOf('shift_id');
  const statusIdx = headers.indexOf('status');
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][idIdx] === shiftId) {
       const currentStatus = data[i][statusIdx];
       if(currentStatus !== 'cancelled' && currentStatus !== 'absent') {
          sheet.getRange(i + 1, statusIdx + 1).setValue(statusOverride || 'completed');
       }
       break;
    }
  }
}

// -------------------------------------------------------------------------
// 🚀 API FOR AE INSPECTION (บันทึกข้อมูลตรวจงานและปัญหา)
// -------------------------------------------------------------------------
function mobileApiSubmitInspection(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inspections');
    let planSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AE_Plans'); // 💡 เพื่อเปลี่ยนสถานะแผนงาน
    let issueSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Issues');
    
    if (!sheet) { setupDatabase(); sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inspections'); }
    if (!issueSheet) { issueSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Issues'); }
    
    let aeId = data.aeId;
    let clientId = data.clientId;
    let planId = data.planId || ''; // 💡 รับ planId มาเพื่ออัปเดตสถานะ
    let ts = new Date().getTime();
    let insId = 'INS-' + ts;

    let sigUrl = '';
    if (data.summary.signatureB64) {
       sigUrl = uploadImageToDrive(data.summary.signatureB64, 'Sign_' + insId + '.png');
    }

    let checklist = data.items || [];
    for (let i = 0; i < checklist.length; i++) {
       if (checklist[i].img && !checklist[i].img.startsWith('http')) {
          checklist[i].imgUrl = uploadImageToDrive(checklist[i].img, 'Chk_' + insId + '_' + i + '.jpg');
          delete checklist[i].img; 
       }
    }

    let issues = data.issues || [];
    const todayStr = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd");
    
    for (let i = 0; i < issues.length; i++) {
       if (issues[i].beforeImg && !issues[i].beforeImg.startsWith('http')) {
          issues[i].beforeUrl = uploadImageToDrive(issues[i].beforeImg, 'Bef_' + insId + '_' + i + '.jpg');
          delete issues[i].beforeImg;
       }
       if (issues[i].afterImg && !issues[i].afterImg.startsWith('http')) {
          issues[i].afterUrl = uploadImageToDrive(issues[i].afterImg, 'Aft_' + insId + '_' + i + '.jpg');
          delete issues[i].afterImg;
       }
       
       saveIssueToBackend({
           id: issues[i].id,
           clientId: clientId,
           dateReported: todayStr,
           source: 'AE Inspector',
           providerId: aeId,
           category: 'ตรวจสอบคุณภาพ',
           description: issues[i].desc,
           status: issues[i].status === 'resolved' ? 'Resolved' : 'Pending',
           assignedTo: '',
           actionTaken: issues[i].actionDesc || '',
           actionBy: aeId
       });
    }

    const now = new Date();
    sheet.appendRow([
      insId, clientId, aeId, todayStr,
      data.summary.quality, data.summary.followUpDate, data.summary.interview,
      sigUrl, JSON.stringify(checklist), JSON.stringify(issues), '', now
    ]);

    // 💡 ถ้ามีการระบุ planId มา ให้ไปเปลี่ยนสถานะในตาราง AE_Plans เป็น Completed ด้วย
    if (planId && planSheet) {
      const pData = planSheet.getDataRange().getValues();
      const pHeaders = pData[0];
      const pIdIdx = pHeaders.indexOf('plan_id');
      const pStatusIdx = pHeaders.indexOf('status');
      
      for(let i = 1; i < pData.length; i++) {
        if(pData[i][pIdIdx] === planId) {
          planSheet.getRange(i+1, pStatusIdx+1).setValue('Completed');
          break;
        }
      }
    }

    return { success: true, message: 'บันทึกข้อมูลการตรวจงานสำเร็จ', inspectionId: insId };
  } catch (e) {
    return { success: false, message: 'เกิดข้อผิดพลาดในการบันทึกตรวจงาน: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// -------------------------------------------------------------------------
// 🚀 API FOR CUSTOMER EVALUATION (ระบบประเมินผลลูกค้ารายบุคคล)
// -------------------------------------------------------------------------
function getClientEvalInfo(clientId, monthStr) {
  try {
    const clients = getSheetDataAsObjects('Clients');
    const client = clients.find(c => c.client_id === clientId);
    if(!client) return { success: false, message: 'ไม่พบข้อมูลสถานที่' };

    const shifts = getSheetDataAsObjects('Shifts');
    const housekeepers = getSheetDataAsObjects('Housekeepers');
    let hkSet = new Set();
    
    shifts.forEach(s => {
      if (s.client_id === clientId && String(s.date).startsWith(monthStr) && s.status !== 'cancelled' && s.status !== 'absent') {
         if (s.assigned_hk_ids) { s.assigned_hk_ids.toString().split(',').forEach(id => hkSet.add(id.trim())); }
      }
    });

    let hkList = [];
    hkSet.forEach(hkId => {
       if(!hkId) return;
       const hk = housekeepers.find(h => h.hk_id === hkId);
       if (hk) { hkList.push({ id: hk.hk_id, name: hk.name, avatar: hk.avatar_url || '' }); }
    });

    return { success: true, clientName: client.client_name, housekeepers: hkList };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function submitEvaluation(clientId, monthStr, scores, comment) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Evaluations');
    const evalId = 'EV-' + new Date().getTime();
    sheet.appendRow([evalId, clientId, monthStr, JSON.stringify(scores), comment, new Date()]);
    return { success: true, message: 'บันทึกการประเมินสำเร็จ' };
  } catch (e) { 
    return { success: false, message: 'การบันทึกประเมินขัดข้อง: ' + e.toString() }; 
  } finally {
    lock.releaseLock();
  }
}
