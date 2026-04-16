/**
 * ==========================================
 * Mari Services - Workforce Scheduling Backend
 * ==========================================
 */

/**
 * ฟังก์ชันนี้สำหรับรันครั้งแรกเท่านั้น! 
 * มันจะช่วยสร้าง Sheets พื้นฐาน และแท็บ ChangeLog เพื่อเก็บประวัติการแก้ไข
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const sheetsConfig = {
    'Housekeepers': ['hk_id', 'name', 'nickname', 'phone', 'line_id', 'status', 'job_type', 'special_skills', 'zones', 'max_hours_week', 'avatar_url', 'color_hex', 'start_date', 'end_date', 'created_at'],
    'Clients': ['client_id', 'client_name', 'address', 'district', 'province', 'type', 'contact_person', 'phone', 'contract_hours', 'required_hk_per_day', 'color_hex', 'status', 'created_at'],
    'Shifts': ['shift_id', 'client_id', 'date', 'start_time', 'end_time', 'assigned_hk_ids', 'status', 'recurring_group_id', 'notes', 'created_by', 'updated_at'],
    'Users': ['email', 'name', 'role', 'is_active'],
    'ChangeLog': ['log_id', 'timestamp', 'user_email', 'action', 'table_name', 'record_id', 'old_data', 'new_data']
  };

  for (const [sheetName, headers] of Object.entries(sheetsConfig)) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    
    // เขียนทับเฉพาะ Header แถวที่ 1 เท่านั้น ข้อมูลเก่าปลอดภัย
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#3d5a6c'); 
    headerRange.setFontColor('white');
    sheet.setFrozenRows(1);
  }

  SpreadsheetApp.getUi().alert('✅ โครงสร้างฐานข้อมูลอัปเดตเรียบร้อย!');
}

/**
 * ฟังก์ชันหลักที่ใช้เปิด Web App
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('Mari Services - Schedule Board')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * ==========================================
 * 🚀 API สำหรับสื่อสารกับ Frontend (Web App)
 * ==========================================
 */

// ระบบเช็ค Login ด้วยบัญชี Google
function verifyUserLogin() {
  const email = Session.getActiveUser().getEmail(); 
  
  if (!email) {
    return { status: 'error', message: 'ไม่สามารถดึงอีเมลได้ กรุณาล็อกอินด้วยบัญชี Google ของท่าน' };
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  if (!sheet) return { status: 'error', message: 'ไม่พบตารางข้อมูล Users ในฐานข้อมูล' };

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const emailIdx = headers.indexOf('email');
  const nameIdx = headers.indexOf('name');
  const roleIdx = headers.indexOf('role');
  const activeIdx = headers.indexOf('is_active');

  // 💡 กรณีใช้งานครั้งแรกสุด (ระบบว่างเปล่า) ให้คนแรกที่เข้าเป็น "Admin / Supervisor" ทันที
  if (data.length === 1 || (data.length === 2 && data[1][0] === '')) {
    if(data.length === 2 && data[1][0] === '') sheet.deleteRow(2); // ลบแถวว่างถ้ามี
    sheet.appendRow([email, 'System Admin', 'Admin / Supervisor', true]);
    return { status: 'success', user: { email: email, name: 'System Admin', role: 'Admin / Supervisor' } };
  }

  // ค้นหาอีเมลในระบบ
  for (let i = 1; i < data.length; i++) {
    if (data[i][emailIdx] === email) {
      const isActive = data[i][activeIdx];
      if (isActive !== true && String(isActive).toLowerCase() !== 'true' && isActive !== 'Active') {
          return { status: 'inactive', message: 'บัญชีของคุณถูกระงับการใช้งาน กรุณาติดต่อ Admin' };
      }
      return {
        status: 'success',
        user: {
          email: email,
          name: data[i][nameIdx],
          role: data[i][roleIdx]
        }
      };
    }
  }

  return { status: 'unauthorized', message: `คุณไม่มีสิทธิ์เข้าถึงระบบนี้ (${email}) กรุณาติดต่อ Admin เพื่อเพิ่มสิทธิ์` };
}

function uploadImageToDrive(base64Data, fileName) {
  try {
    var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), MimeType.JPEG, fileName);
    var FOLDER_ID = '1UGFSugaAD8CWC7oWWPMhuoYc3puvMnZA'; 
    var folder = DriveApp.getFolderById(FOLDER_ID);
    var file = folder.createFile(blob);
    
    try {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (shareError) {
        try { file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); } 
        catch (domainError) { console.log("Policy Block: " + domainError); }
    }
    
    var url = "https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w500";
    return url;
  } catch (e) {
    throw new Error(e.toString());
  }
}

function getAppData() {
  return {
    clients: getSheetDataAsObjects('Clients'),
    housekeepers: getSheetDataAsObjects('Housekeepers'),
    shifts: getSheetDataAsObjects('Shifts'),
    users: getSheetDataAsObjects('Users')
  };
}

// ==========================================
// 🚀 ฟังก์ชันบันทึกการเปลี่ยนแปลง (Change Log)
// ==========================================
function logChange(action, tableName, recordId, oldData, newData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ChangeLog');
  if (!sheet) return;
  const email = Session.getActiveUser().getEmail() || 'Unknown';
  sheet.appendRow([
    'LOG-' + new Date().getTime(),
    new Date(),
    email,
    action, // CREATE, UPDATE, DELETE
    tableName,
    recordId,
    oldData ? JSON.stringify(oldData) : '',
    newData ? JSON.stringify(newData) : ''
  ]);
}

// ==========================================
// 🚀 ฟังก์ชันตรวจสอบความขัดแย้ง (Conflict Checking)
// ==========================================
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


// ==========================================
// 🚀 ฟังก์ชันบันทึกและอัปเดตข้อมูล (SAVE)
// ==========================================

function saveClientToBackend(clientData) {
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

      sheet.getRange(rowNum, headers.indexOf('client_name') + 1).setValue(clientData.name || '');
      sheet.getRange(rowNum, headers.indexOf('address') + 1).setValue(clientData.address || '');
      sheet.getRange(rowNum, headers.indexOf('district') + 1).setValue(clientData.district || '');
      sheet.getRange(rowNum, headers.indexOf('province') + 1).setValue(clientData.province || '');
      sheet.getRange(rowNum, headers.indexOf('type') + 1).setValue(clientData.type || 'B2B');
      sheet.getRange(rowNum, headers.indexOf('contact_person') + 1).setValue(clientData.contact || '');
      sheet.getRange(rowNum, headers.indexOf('phone') + 1).setValue(clientData.phone || '');
      sheet.getRange(rowNum, headers.indexOf('contract_hours') + 1).setValue(clientData.contractHours || '');
      sheet.getRange(rowNum, headers.indexOf('required_hk_per_day') + 1).setValue(clientData.reqStaff || 1);
      sheet.getRange(rowNum, headers.indexOf('color_hex') + 1).setValue(clientData.color || '#e2e8f0');
      sheet.getRange(rowNum, headers.indexOf('status') + 1).setValue(clientData.status || 'Active');
      isFound = true;
      
      logChange('UPDATE', 'Clients', clientData.id, oldDataObj, clientData);
      break;
    }
  }

  if (!isFound) {
    sheet.appendRow([
      clientData.id || 'CL-' + new Date().getTime(),
      clientData.name || '',
      clientData.address || '',
      clientData.district || '',
      clientData.province || '',
      clientData.type || 'B2B',
      clientData.contact || '',
      clientData.phone || '',
      clientData.contractHours || '',
      clientData.reqStaff || 1,
      clientData.color || '#e2e8f0',
      clientData.status || 'Active',
      new Date()
    ]);
    logChange('CREATE', 'Clients', clientData.id, null, clientData);
  }
  return { success: true, message: 'บันทึกข้อมูลไซต์งานสำเร็จ' };
}

function saveStaffToBackend(staffData) {
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

      sheet.getRange(rowNum, headers.indexOf('name') + 1).setValue(staffData.name || '');
      sheet.getRange(rowNum, headers.indexOf('nickname') + 1).setValue(staffData.nickname || '');
      sheet.getRange(rowNum, headers.indexOf('phone') + 1).setValue(staffData.phone || '');
      sheet.getRange(rowNum, headers.indexOf('line_id') + 1).setValue(staffData.lineId || '');
      sheet.getRange(rowNum, headers.indexOf('status') + 1).setValue(staffData.status || 'Active');
      sheet.getRange(rowNum, headers.indexOf('job_type') + 1).setValue(staffData.type || 'Full-time');
      sheet.getRange(rowNum, headers.indexOf('special_skills') + 1).setValue(staffData.skills || '');
      sheet.getRange(rowNum, headers.indexOf('zones') + 1).setValue(staffData.zones || '');
      sheet.getRange(rowNum, headers.indexOf('max_hours_week') + 1).setValue(staffData.maxHoursWeek || 48);
      sheet.getRange(rowNum, headers.indexOf('start_date') + 1).setValue(staffData.startDate || '');
      sheet.getRange(rowNum, headers.indexOf('end_date') + 1).setValue(staffData.endDate || '');
      sheet.getRange(rowNum, headers.indexOf('avatar_url') + 1).setValue(staffData.avatar || '');
      sheet.getRange(rowNum, headers.indexOf('color_hex') + 1).setValue(staffData.color || '#3b82f6');
      isFound = true;

      logChange('UPDATE', 'Housekeepers', staffData.id, oldDataObj, staffData);
      break;
    }
  }

  if (!isFound) {
    sheet.appendRow([
      staffData.id || 'HK-' + new Date().getTime(),
      staffData.name || '',
      staffData.nickname || '',
      staffData.phone || '',
      staffData.lineId || '',
      staffData.status || 'Active',
      staffData.type || 'Full-time',
      staffData.skills || '',
      staffData.zones || '',
      staffData.maxHoursWeek || 48,
      staffData.avatar || '',
      staffData.color || '#3b82f6',
      staffData.startDate || '', 
      staffData.endDate || '', 
      new Date()
    ]);
    logChange('CREATE', 'Housekeepers', staffData.id, null, staffData);
  }
  return { success: true, message: 'บันทึกข้อมูลพนักงานสำเร็จ' };
}

function saveUserToBackend(userData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  if (!sheet) return { success: false, message: 'ไม่พบ Sheet: Users' };

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const emailIdx = headers.indexOf('email');
  
  let isFound = false;

  for (let i = 1; i < data.length; i++) {
    if (data[i][emailIdx] === userData.email || data[i][emailIdx] === userData.id) {
      const rowNum = i + 1;
      sheet.getRange(rowNum, headers.indexOf('name') + 1).setValue(userData.name || '');
      sheet.getRange(rowNum, headers.indexOf('role') + 1).setValue(userData.role || 'Viewer');
      const isActive = (userData.status === 'Active') ? true : false;
      sheet.getRange(rowNum, headers.indexOf('is_active') + 1).setValue(isActive);
      isFound = true;
      break;
    }
  }

  if (!isFound) {
    const isActive = (userData.status === 'Active') ? true : false;
    sheet.appendRow([
      userData.email,
      userData.name || '',
      userData.role || 'Viewer',
      isActive
    ]);
  }
  
  return { success: true, message: 'บันทึกข้อมูลผู้ใช้งานสำเร็จ' };
}

function saveMultipleShiftsToBackend(shiftsArray) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Shifts');
  if (!sheet) return { success: false, message: 'ไม่พบ Sheet: Shifts' };

  const warnings = checkShiftConflicts(shiftsArray);

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const shiftIdIdx = headers.indexOf('shift_id');
  const email = Session.getActiveUser().getEmail();
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

        sheet.getRange(rowNum, headers.indexOf('client_id') + 1).setValue(shiftData.clientId);
        sheet.getRange(rowNum, headers.indexOf('date') + 1).setValue(shiftData.date);
        sheet.getRange(rowNum, headers.indexOf('start_time') + 1).setValue(shiftData.start);
        sheet.getRange(rowNum, headers.indexOf('end_time') + 1).setValue(shiftData.end);
        sheet.getRange(rowNum, headers.indexOf('assigned_hk_ids') + 1).setValue(hkString);
        sheet.getRange(rowNum, headers.indexOf('status') + 1).setValue(shiftData.status);
        sheet.getRange(rowNum, headers.indexOf('notes') + 1).setValue(notesStr);
        sheet.getRange(rowNum, headers.indexOf('updated_at') + 1).setValue(now);
        isFound = true;
        
        logChange('UPDATE', 'Shifts', shiftData.id, oldDataObj, shiftData);
        break; 
      }
    }

    if (!isFound) {
       let newRow = new Array(headers.length).fill('');
       newRow[headers.indexOf('shift_id')] = shiftData.id;
       newRow[headers.indexOf('client_id')] = shiftData.clientId;
       newRow[headers.indexOf('date')] = shiftData.date;
       newRow[headers.indexOf('start_time')] = shiftData.start;
       newRow[headers.indexOf('end_time')] = shiftData.end;
       newRow[headers.indexOf('assigned_hk_ids')] = hkString;
       newRow[headers.indexOf('status')] = shiftData.status;
       newRow[headers.indexOf('recurring_group_id')] = groupIdStr;
       newRow[headers.indexOf('notes')] = notesStr;
       newRow[headers.indexOf('created_by')] = email;
       newRow[headers.indexOf('updated_at')] = now;
       
       newRows.push(newRow);
       logChange('CREATE', 'Shifts', shiftData.id, null, shiftData);
    }
  });

  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, headers.length).setValues(newRows);
  }

  return { 
    success: true, 
    message: `บันทึกตารางงานสำเร็จ ${shiftsArray.length} รายการ`,
    warnings: warnings 
  };
}

function updateShiftDragAndDrop(shiftId, targetClientId, targetDateStr) {
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

      sheet.getRange(rowNum, headers.indexOf('client_id') + 1).setValue(targetClientId);
      sheet.getRange(rowNum, headers.indexOf('date') + 1).setValue(targetDateStr);
      sheet.getRange(rowNum, headers.indexOf('updated_at') + 1).setValue(new Date());
      
      logChange('UPDATE_DRAG', 'Shifts', shiftId, oldDataObj, {clientId: targetClientId, date: targetDateStr});
      return { success: true, message: 'อัปเดตตำแหน่งสำเร็จ' };
    }
  }
  return { success: false, message: 'ไม่พบ Shift ID นี้ในระบบ' };
}

function deleteShiftToBackend(shiftId, deleteType = 'single', groupId = null) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Shifts');
  if (!sheet) return { success: false, message: 'ไม่พบ Sheet: Shifts' };

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const shiftIdIdx = headers.indexOf('shift_id');
  const groupIdIdx = headers.indexOf('recurring_group_id');
  
  let deletedCount = 0;

  for (let i = data.length - 1; i >= 1; i--) {
    let shouldDelete = false;
    
    if (deleteType === 'group' && groupId && data[i][groupIdIdx] === groupId) {
      shouldDelete = true;
    } else if (data[i][shiftIdIdx] === shiftId) {
      shouldDelete = true;
    }

    if (shouldDelete) {
      let oldDataObj = {};
      for(let j=0; j<headers.length; j++) { oldDataObj[headers[j]] = data[i][j]; }
      
      sheet.deleteRow(i + 1);
      logChange('DELETE', 'Shifts', data[i][shiftIdIdx], oldDataObj, null);
      deletedCount++;

      if (deleteType === 'single') break;
    }
  }

  if (deletedCount > 0) {
    return { success: true, message: `ลบตารางงานสำเร็จ ${deletedCount} รายการ` };
  }
  
  return { success: false, message: 'ไม่พบข้อมูลที่ต้องการลบ' };
}

function deleteStaffToBackend(staffId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Housekeepers');
  if (!sheet) return { success: false, message: 'ไม่พบ Sheet: Housekeepers' };

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx = headers.indexOf('hk_id');
  
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][idIdx] === staffId) {
      let oldDataObj = {};
      for(let j=0; j<headers.length; j++) { oldDataObj[headers[j]] = data[i][j]; }
      
      sheet.deleteRow(i + 1);
      logChange('DELETE', 'Housekeepers', staffId, oldDataObj, null);
      return { success: true, message: 'ลบข้อมูลพนักงานสำเร็จ' };
    }
  }
  return { success: false, message: 'ไม่พบข้อมูลที่ต้องการลบ' };
}

function deleteClientToBackend(clientId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clients');
  if (!sheet) return { success: false, message: 'ไม่พบ Sheet: Clients' };

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx = headers.indexOf('client_id');
  
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][idIdx] === clientId) {
      let oldDataObj = {};
      for(let j=0; j<headers.length; j++) { oldDataObj[headers[j]] = data[i][j]; }

      sheet.deleteRow(i + 1);
      logChange('DELETE', 'Clients', clientId, oldDataObj, null);
      return { success: true, message: 'ลบข้อมูลไซต์งานสำเร็จ' };
    }
  }
  return { success: false, message: 'ไม่พบข้อมูลที่ต้องการลบ' };
}

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
        if (headers[j] === 'date' || headers[j] === 'start_date' || headers[j] === 'end_date') {
          const year = value.getFullYear();
          const month = String(value.getMonth() + 1).padStart(2, '0');
          const day = String(value.getDate()).padStart(2, '0');
          value = `${year}-${month}-${day}`;
        }
        else if (headers[j] === 'start_time' || headers[j] === 'end_time') {
           const hours = String(value.getHours()).padStart(2, '0');
           const minutes = String(value.getMinutes()).padStart(2, '0');
           value = `${hours}:${minutes}`;
        }
        else {
           value = value.toISOString();
        }
      }
      
      if (headers[j] === 'avatar_url') obj['avatar'] = value;
      if (headers[j] === 'color_hex') obj['color'] = value;
      
      obj[headers[j]] = value;
    }
    result.push(obj);
  }
  return result;
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getAddressDataFromSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config province');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return []; 
  const result = [];
  for (let i = 1; i < data.length; i++) {
    result.push({ postCode: data[i][0] || '', ProvinceThai: data[i][1] || '', DistrictThai: data[i][2] || '', TambonThai: data[i][3] || '' });
  }
  return result;
}

// ฟังก์ชันสำหรับดึงประวัติของกะงานจากชีต ChangeLog
function getShiftHistory(shiftId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ChangeLog');
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0];
  const recordIdIdx = headers.indexOf('record_id');
  const result = [];
  
  // ดึงข้อมูลจากล่างขึ้นบน (เพื่อให้ประวัติล่าสุดอยู่บนสุด)
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][recordIdIdx] === shiftId) {
      result.push({
        timestamp: data[i][headers.indexOf('timestamp')],
        user_email: data[i][headers.indexOf('user_email')],
        action: data[i][headers.indexOf('action')]
      });
    }
  }
  return result;
}

/**
 * 🚀 ฟังก์ชันสำหรับ Export รายงานตารางงานไปสร้างเป็น Google Sheets
 */
function exportToGoogleSheets(shiftsData) {
  try {
    // 1. สร้างไฟล์ Sheet ใหม่
    const timeStamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
    const ss = SpreadsheetApp.create('MARI_Schedule_Export_' + timeStamp);
    const sheet = ss.getActiveSheet();
    
    // 2. สร้างส่วนหัวของตาราง (Headers)
    const headers = ['รหัสกะงาน', 'วันที่ปฏิบัติงาน', 'เวลาเริ่ม', 'เวลาสิ้นสุด', 'ลูกค้า / สถานที่', 'รายชื่อพนักงาน', 'สถานะ', 'หมายเหตุ'];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#3d5a6c').setFontColor('white');
    
    // 3. เตรียมข้อมูล
    const clients = getSheetDataAsObjects('Clients');
    const housekeepers = getSheetDataAsObjects('Housekeepers');
    
    const rows = shiftsData.map(shift => {
      const client = clients.find(c => c.client_id === shift.clientId) || {client_name: shift.clientId};
      const hks = shift.hks ? shift.hks.map(hkId => {
        const h = housekeepers.find(x => x.hk_id === hkId);
        return h ? h.name : hkId;
      }).join(', ') : '';
      
      return [
        shift.id,
        shift.date,
        shift.start,
        shift.end,
        client.client_name,
        hks,
        shift.status,
        shift.notes || ''
      ];
    });
    
    // 4. เขียนข้อมูลลงชีต และจัดความกว้างอัตโนมัติ
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }
    sheet.autoResizeColumns(1, headers.length);
    
    // คืนค่า URL ให้หน้าบ้าน เพื่อเปิดชีตในหน้าต่างใหม่
    return ss.getUrl();
    
  } catch (e) {
    throw new Error('เกิดข้อผิดพลาดในการสร้าง Sheet: ' + e.toString());
  }
}

