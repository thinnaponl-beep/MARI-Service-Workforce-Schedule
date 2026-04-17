/**
 * ==========================================
 * Mari Services - Workforce Scheduling Backend
 * ==========================================
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
    
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#3d5a6c'); 
    headerRange.setFontColor('white');
    sheet.setFrozenRows(1);
  }

  SpreadsheetApp.getUi().alert('✅ โครงสร้างฐานข้อมูลอัปเดตเรียบร้อย!');
}

function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('Mari Services - Schedule Board')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function verifyUserLogin() {
  const email = Session.getActiveUser().getEmail(); 
  if (!email) return { status: 'error', message: 'ไม่สามารถดึงอีเมลได้ กรุณาล็อกอินด้วยบัญชี Google ของท่าน' };

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  if (!sheet) return { status: 'error', message: 'ไม่พบตารางข้อมูล Users ในฐานข้อมูล' };

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const emailIdx = headers.indexOf('email');
  const nameIdx = headers.indexOf('name');
  const roleIdx = headers.indexOf('role');
  const activeIdx = headers.indexOf('is_active');

  if (data.length === 1 || (data.length === 2 && data[1][0] === '')) {
    if(data.length === 2 && data[1][0] === '') sheet.deleteRow(2);
    sheet.appendRow([email, 'System Admin', 'Admin / Supervisor', true]);
    return { status: 'success', user: { email: email, name: 'System Admin', role: 'Admin / Supervisor' } };
  }

  for (let i = 1; i < data.length; i++) {
    if (data[i][emailIdx] === email) {
      const isActive = data[i][activeIdx];
      if (isActive !== true && String(isActive).toLowerCase() !== 'true' && isActive !== 'Active') {
          return { status: 'inactive', message: 'บัญชีของคุณถูกระงับการใช้งาน กรุณาติดต่อ Admin' };
      }
      return { status: 'success', user: { email: email, name: data[i][nameIdx], role: data[i][roleIdx] } };
    }
  }

  return { status: 'unauthorized', message: `คุณไม่มีสิทธิ์เข้าถึงระบบนี้ (${email}) กรุณาติดต่อ Admin เพื่อเพิ่มสิทธิ์` };
}

// 💡 อัปเดตโดยศักดิ์ชัย: แก้ปัญหาภาพไม่แสดงในมือถือ/iFrame
function uploadImageToDrive(base64Data, fileName) {
  try {
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
    
    // 💡 THE ULTIMATE FIX: เปลี่ยนมาใช้ URL lh3.googleusercontent.com (ทะลุการบล็อก 100%)
    var imageUrl = "https://lh3.googleusercontent.com/d/" + fileId;
    
    return imageUrl;
    
  } catch (e) { 
    throw new Error("Upload failed: " + e.toString()); 
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

function logChange(action, tableName, recordId, oldData, newData, actionBy) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ChangeLog');
  if (!sheet) return;
  const email = actionBy || Session.getActiveUser().getEmail() || 'Unknown';
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

      let updatedRow = [...data[i]];
      updatedRow[headers.indexOf('client_name')] = clientData.name || '';
      updatedRow[headers.indexOf('address')] = clientData.address || '';
      updatedRow[headers.indexOf('district')] = clientData.district || '';
      updatedRow[headers.indexOf('province')] = clientData.province || '';
      updatedRow[headers.indexOf('type')] = clientData.type || 'B2B';
      updatedRow[headers.indexOf('contact_person')] = clientData.contact || '';
      updatedRow[headers.indexOf('phone')] = clientData.phone || '';
      updatedRow[headers.indexOf('contract_hours')] = clientData.contractHours || '';
      updatedRow[headers.indexOf('required_hk_per_day')] = clientData.reqStaff || 1;
      updatedRow[headers.indexOf('color_hex')] = clientData.color || '#e2e8f0';
      updatedRow[headers.indexOf('status')] = clientData.status || 'Active';

      sheet.getRange(rowNum, 1, 1, headers.length).setValues([updatedRow]);
      isFound = true;
      
      logChange('UPDATE', 'Clients', clientData.id, oldDataObj, clientData, clientData.actionBy);
      break;
    }
  }

  if (!isFound) {
    sheet.appendRow([
      clientData.id || 'CL-' + new Date().getTime(),
      clientData.name || '', clientData.address || '', clientData.district || '', clientData.province || '',
      clientData.type || 'B2B', clientData.contact || '', clientData.phone || '', clientData.contractHours || '',
      clientData.reqStaff || 1, clientData.color || '#e2e8f0', clientData.status || 'Active', new Date()
    ]);
    logChange('CREATE', 'Clients', clientData.id, null, clientData, clientData.actionBy);
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

      let updatedRow = [...data[i]];
      updatedRow[headers.indexOf('name')] = staffData.name || '';
      updatedRow[headers.indexOf('nickname')] = staffData.nickname || '';
      updatedRow[headers.indexOf('phone')] = staffData.phone || '';
      updatedRow[headers.indexOf('line_id')] = staffData.lineId || '';
      updatedRow[headers.indexOf('status')] = staffData.status || 'Active';
      updatedRow[headers.indexOf('job_type')] = staffData.type || 'Full-time';
      updatedRow[headers.indexOf('special_skills')] = staffData.skills || '';
      updatedRow[headers.indexOf('zones')] = staffData.zones || '';
      updatedRow[headers.indexOf('max_hours_week')] = staffData.maxHoursWeek || 48;
      updatedRow[headers.indexOf('start_date')] = staffData.startDate || '';
      updatedRow[headers.indexOf('end_date')] = staffData.endDate || '';
      updatedRow[headers.indexOf('avatar_url')] = staffData.avatar || '';
      updatedRow[headers.indexOf('color_hex')] = staffData.color || '#3b82f6';

      sheet.getRange(rowNum, 1, 1, headers.length).setValues([updatedRow]);
      isFound = true;

      logChange('UPDATE', 'Housekeepers', staffData.id, oldDataObj, staffData, staffData.actionBy);
      break;
    }
  }

  if (!isFound) {
    sheet.appendRow([
      staffData.id || 'HK-' + new Date().getTime(),
      staffData.name || '', staffData.nickname || '', staffData.phone || '', staffData.lineId || '',
      staffData.status || 'Active', staffData.type || 'Full-time', staffData.skills || '', staffData.zones || '',
      staffData.maxHoursWeek || 48, staffData.avatar || '', staffData.color || '#3b82f6',
      staffData.startDate || '', staffData.endDate || '', new Date()
    ]);
    logChange('CREATE', 'Housekeepers', staffData.id, null, staffData, staffData.actionBy);
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
      let updatedRow = [...data[i]];
      updatedRow[headers.indexOf('name')] = userData.name || '';
      updatedRow[headers.indexOf('role')] = userData.role || 'Viewer';
      updatedRow[headers.indexOf('is_active')] = (userData.status === 'Active') ? true : false;

      sheet.getRange(rowNum, 1, 1, headers.length).setValues([updatedRow]);
      isFound = true;
      break;
    }
  }

  if (!isFound) {
    const isActive = (userData.status === 'Active') ? true : false;
    sheet.appendRow([ userData.email, userData.name || '', userData.role || 'Viewer', isActive ]);
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
        updatedRow[headers.indexOf('client_id')] = shiftData.clientId;
        updatedRow[headers.indexOf('date')] = shiftData.date;
        updatedRow[headers.indexOf('start_time')] = shiftData.start;
        updatedRow[headers.indexOf('end_time')] = shiftData.end;
        updatedRow[headers.indexOf('assigned_hk_ids')] = hkString;
        updatedRow[headers.indexOf('status')] = shiftData.status;
        updatedRow[headers.indexOf('notes')] = notesStr;
        updatedRow[headers.indexOf('updated_at')] = now;

        sheet.getRange(rowNum, 1, 1, headers.length).setValues([updatedRow]);
        isFound = true;
        
        let logAction = shiftData.actionType || 'UPDATE';
        logChange(logAction, 'Shifts', shiftData.id, oldDataObj, shiftData, shiftData.actionBy);
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
       newRow[headers.indexOf('created_by')] = shiftData.actionBy || 'Unknown';
       newRow[headers.indexOf('updated_at')] = now;
       
       newRows.push(newRow);
       logChange('CREATE', 'Shifts', shiftData.id, null, shiftData, shiftData.actionBy);
    }
  });

  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, headers.length).setValues(newRows);
  }

  return { success: true, message: `บันทึกตารางงานสำเร็จ ${shiftsArray.length} รายการ`, warnings: warnings };
}

function updateShiftDragAndDrop(shiftId, targetClientId, targetDateStr, actionBy) {
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
      
      logChange('UPDATE_DRAG', 'Shifts', shiftId, oldDataObj, {clientId: targetClientId, date: targetDateStr}, actionBy);
      return { success: true, message: 'อัปเดตตำแหน่งสำเร็จ' };
    }
  }
  return { success: false, message: 'ไม่พบ Shift ID นี้ในระบบ' };
}

function deleteShiftToBackend(shiftId, deleteType = 'single', groupId = null, actionBy) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Shifts');
  if (!sheet) return { success: false, message: 'ไม่พบ Sheet: Shifts' };

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const shiftIdIdx = headers.indexOf('shift_id');
  const groupIdIdx = headers.indexOf('recurring_group_id');
  let deletedCount = 0;

  for (let i = data.length - 1; i >= 1; i--) {
    let shouldDelete = false;
    if (deleteType === 'group' && groupId && data[i][groupIdIdx] === groupId) shouldDelete = true;
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
}

function deleteStaffToBackend(staffId, actionBy) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Housekeepers');
  if (!sheet) return { success: false, message: 'ไม่พบ Sheet: Housekeepers' };
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][headers.indexOf('hk_id')] === staffId) {
      let oldDataObj = {};
      for(let j=0; j<headers.length; j++) { oldDataObj[headers[j]] = data[i][j]; }
      sheet.deleteRow(i + 1);
      logChange('DELETE', 'Housekeepers', staffId, oldDataObj, null, actionBy);
      return { success: true, message: 'ลบข้อมูลพนักงานสำเร็จ' };
    }
  }
  return { success: false, message: 'ไม่พบข้อมูลที่ต้องการลบ' };
}

function deleteClientToBackend(clientId, actionBy) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clients');
  if (!sheet) return { success: false, message: 'ไม่พบ Sheet: Clients' };
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][headers.indexOf('client_id')] === clientId) {
      let oldDataObj = {};
      for(let j=0; j<headers.length; j++) { oldDataObj[headers[j]] = data[i][j]; }
      sheet.deleteRow(i + 1);
      logChange('DELETE', 'Clients', clientId, oldDataObj, null, actionBy);
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
    if (String(data[i][headers.indexOf('table_name')]) === String(tableName) && 
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

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
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
